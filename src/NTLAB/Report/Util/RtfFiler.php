<?php

/*
 * The MIT License
 *
 * Copyright (c) 2014 Toha <tohenk@yahoo.com>
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of
 * this software and associated documentation files (the "Software"), to deal in
 * the Software without restriction, including without limitation the rights to
 * use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies
 * of the Software, and to permit persons to whom the Software is furnished to do
 * so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */

namespace NTLAB\Report\Util;

use NTLAB\Script\Core\Script;
use NTLAB\RtfTree\Node\Tree;
use NTLAB\RtfTree\Node\Node;
use NTLAB\Report\Util\Extractor\Extractor;
use NTLAB\Report\Util\Extractor\Paragraph as ParagraphExtractor;
use NTLAB\Report\Util\Extractor\Table as TableExtractor;

class RtfFiler
{
    /**
     * @var \NTLAB\Report\Util\RtfFiler
     */
    protected static $instance = null;

    /**
     * @var \NTLAB\Script\Core\Script
     */
    protected $script = null;

    /**
     * Get the instance.
     *
     * @return \NTLAB\Report\Util\RtfFiler
     */
    public static function getInstance()
    {
        if (null == self::$instance) {
            self::$instance = new self();
        }

        return self::$instance;
    }

    /**
     * Get script object.
     *
     * @return \NTLAB\Script\Core\Script
     */
    public function getScript()
    {
        if (null == $this->script) {
            $this->script = new Script();
        }

        return $this->script;
    }

    /**
     * Set the script object.
     *
     * @param \NTLAB\Script\Core\Script $script  The script object
     * @return \NTLAB\Report\Util\RtfFiler
     */
    public function setScript(Script $script)
    {
        $this->script = $script;

        return $this;
    }

    /**
     * Load rtf tree from template.
     *
     * @param string $template  Richtext content
     * @return \NTLAB\RtfTree\Node\Tree
     */
    protected function load($template)
    {
        $tree = new Tree();
        $tree->setIgnoreWhitespace(false);
        if ($tree->loadFromString($template)) {
            return $tree;
        }
        unset($tree);
    }

    /**
     * Build and parse template tree.
     *
     * @param \NTLAB\RtfTree\Node\Tree $tree  Template tree
     * @param mixed $objects  The objects
     * @return \NTLAB\RtfTree\Node\Tree
     */
    protected function doBuild(Tree $tree, $objects)
    {
        $body = ParagraphExtractor::getBody($tree);
        $each = ParagraphExtractor::getRegion($body->getResult(), 'EACH');
        $table = TableExtractor::getTable($body->getResult());

        $result = Extractor::createTree();
        // include header
        Extractor::copyTree($result, $tree, 0, $body->getBeginPos() - 1);
        // process body
        $this->getScript()
            ->setObjects($objects)
            ->each(function(Script $script, RtfFiler $_this) use ($result, $body, $each, $table) {
                $clone = $body->getResult()->cloneTree();
                $_this
                    ->replaceTag($clone, array_merge(array_keys($each->getRegions()), array_keys($table->getRegions())))
                    ->replaceRegion($clone, $each->getRegions(), 'par')
                    ->replaceRegion($clone, $table->getRegions(), 'pard')
                ;
                Extractor::appendTree($result, $clone);
                unset($clone);
            })
        ;
        // include footer
        Extractor::copyTree($result, $tree, $body->getEndPos() + 1, count($tree->getMainGroup()->getChildren()) - 1);

        return $result;
    }

    /**
     * Replace all tags in tree.
     *
     * @param \NTLAB\RtfTree\Node\Tree $tree  The template body tree
     * @param array $ignores  Ignore tags
     * @return \NTLAB\Report\Util\RtfFiler
     */
    public function replaceTag(Tree $tree, $ignores = array())
    {
        $caches = array();
        preg_match_all(Extractor::getTagRegex(), $tree->getText(), $matches, PREG_PATTERN_ORDER);
        for ($i = 0; $i < count($matches[0]); $i++) {
            $tag = $matches[1][$i];
            // check for ignore
            if (in_array($tag, $ignores)) {
                continue;
            }
            $match = $matches[0][$i];
            if (isset($caches[$tag])) {
                $value = $caches[$tag];
            } else {
                if (!$this->getScript()->getVar($value, $tag, $this->script->getContext())) {
                    $value = $this->getScript()->evaluate($tag);
                }
                $caches[$tag] = $value;
            }
            $tree->getMainGroup()->replaceTextEx($match, $value);
        }

        return $this;
    }

    /**
     * Replace region.
     *
     * @param \NTLAB\RtfTree\Node\Tree $tree  Result tree
     * @param array $regions  Region data
     * @param string $eolKey  Eol keyword used to separate each item
     * @return \NTLAB\Report\Util\RtfFiler
     */
    public function replaceRegion(Tree $tree, $regions = array(), $eolKey)
    {
        foreach ($regions as $tag => $data) {
            list($expr, $body) = $data;
            $index = Extractor::findTag($tree, $tag, $size);
            if ($objects = $this->getScript()->evaluate($expr)) {
                $result = Extractor::createTree();
                $this->getScript()
                    ->pushContext()
                    ->setObjects($objects)
                    ->each(function(Script $script, RtfFiler $_this) use ($result, $body, $eolKey) {
                        $clone = $body->cloneTree();
                        $_this->replaceTag($clone);
                        Extractor::appendTree($result, $clone);
                        if ($script->getIterator()->getRecNo() < $script->getIterator()->getRecCount()) {
                            $result->getMainGroup()->appendChild(Node::create(Node::KEYWORD, $eolKey));
                        }
                        unset($clone);
                    }, false)
                    ->popContext()
                ;
                Extractor::insertTree($tree, $result, $index + 1);
                unset($result);
            }
            // remove tag
            $tree->getMainGroup()->removeChildAt($index);
        }

        return $this;
    }

    /**
     * Build the template.
     *
     * @param string $template  Tree template
     * @param mixed $objects  The objects
     * @return string
     */
    public function build($template, $objects)
    {
        if ($tree = $this->load($template)) {
            return $this->doBuild($tree, $objects)
                ->getRtf();
        }
    }
}