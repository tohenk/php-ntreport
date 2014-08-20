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

use NTLAB\Common\Terbilang;
use NTLAB\Script\Core\Script;
use NTLAB\RtfTree\Node\Node;
use NTLAB\RtfTree\Node\Tree;

class RtfFiler
{
    const TAG_SIGN = '%';
    const RTF_PARAGRAPH = 'pard';

    /**
     * @var \NTLAB\Report\Util\RtfFiler
     */
    protected static $instance = null;

    /**
     * @var \NTLAB\Script\Core\Script
     */
    protected $script = null;

    /**
     * @var string
     */
    protected $tagRe = null;

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
     * Get the enclosing tag signs.
     * The result is an array which contains
     * the opening and closing tag sign array(opening, closing).
     *
     * @return array
     */
    public function getTags()
    {
        $stag = substr(self::TAG_SIGN, 0, 1);
        $etag = strlen(self::TAG_SIGN) > 1 ? substr(self::TAG_SIGN, 1, 1) : $stag;

        return array($stag, $etag);
    }

    /**
     * Get tag regular expression pattern.
     *
     * @return string
     */
    public function getTagRegex()
    {
        if (null == $this->tagRe) {
            $tags = $this->getTags();
            $this->tagRe = sprintf('/%1$s([^%1$s]+)%2$s/', $tags[0], $tags[1]);
        }

        return $this->tagRe;
    }

    /**
     * Create a tag.
     *
     * @param string $tag  The tag
     * @return string
     */
    public function createTag($tag)
    {
        $tags = $this->getTags();

        return $tags[0].$tag.$tags[1];
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
     * Create rtf tree.
     *
     * @return \NTLAB\RtfTree\Node\Tree
     */
    protected function createTree()
    {
        $tree = new Tree();
        $tree->addMainGroup();

        return $tree;
    }

    /**
     * Find tag in the tree.
     *
     * @param \NTLAB\Rtf\Node\Tree $tree  The tree 
     * @param string $tag  Tag to find
     * @param int $start  Start position
     * @return int
     */
    protected function findTag(Tree $tree, $tag, $start = null)
    {
        $tag = $this->createTag($tag);
        if (null === $start) {
            $start = 0;
        }
        for ($i = $start; $i < count($tree->getMainGroup()->getChildren()); $i++) {
            if (!($node = $tree->getMainGroup()->getChildAt($i))) {
                continue;
            }
            if (false !== mb_strpos($node->getNodeText(), $tag)) {
                return $i;
            }
        }

        return false;
    }

    /**
     * Extract template for begin and end mark. If begin or ending mark is not
     * found, return all paragraphs within template.
     *
     * @param \NTLAB\Rtf\Node\Tree $tree  Template tree
     * @param int $bpos  Begin mark position found
     * @param int $epos  End mark position found
     * @param string $bmark  Begining mark
     * @param string $emark  Ending mark
     * @return \NTLAB\RtfTree\Node\Tree
     * @throws \InvalidArgumentException
     */
    public function extract(Tree $tree, &$bpos, &$epos, $bmark = 'BEGIN', $emark = 'END')
    {
        if (!($node = $tree->getMainGroup()->selectSingleNode(static::RTF_PARAGRAPH))) {
            throw new \InvalidArgumentException('No paragraph found.');
        }
        $start = $node->getNodeIndex();
        $result = $this->createTree();
        $bpos = $this->findTag($tree, $bmark, $start);
        $epos = $this->findTag($tree, $emark, $bpos);
        // begin and end mark found
        if (false !== $bpos && false !== $epos) {
            $oldBpos = $bpos;
            // ensure first match contain pard
            while (true) {
                if ($bpos === $start) {
                    break;
                }
                $node = $tree->getMainGroup()->getChildAt($bpos);
                if ($node->isEquals(static::RTF_PARAGRAPH)) {
                    break;
                }
                $bpos--;
            }
            for ($i = $bpos; $i <= $epos; $i++) {
                $result->getMainGroup()->appendChild($tree->getMainGroup()->getChildAt($i)->cloneNode());
            }
            // remove begin and end mark
            $result->getMainGroup()->replaceTextEx($this->createTag($bmark), '');
            $result->getMainGroup()->replaceTextEx($this->createTag($emark), '');
        } else {
            $bpos = $start;
            $epos = count($tree->getMainGroup()->getChildren()) - 1;
            // append all paragraph
            for ($i = $bpos; $i <= $epos; $i++) {
                $result->getMainGroup()->appendChild($tree->getMainGroup()->getChildAt($i)->cloneNode());
            }
        }

        return $result;
    }

    /**
     * Find prefix in the matches tags.
     *
     * @param array $matches  The matches
     * @param string $prefix  Tag prefix
     * @return int
     */
    protected function findMatch($matches, $prefix)
    {
        for ($i = 0; $i < count($matches[0]); $i++) {
            $match = $matches[0][$i];
            $tag = $matches[1][$i];
            if (0 === mb_strpos($tag, $prefix)) {
                return $i;
            }
        }

        return false;
    }

    /**
     * Extract region from tree.
     *
     * @param \NTLAB\RtfTree\Node\Tree $tree  The tree to extract
     * @param string $region
     * @return array
     */
    protected function extractRegion(Tree $tree, $region)
    {
        $result = array();
        while (true) {
            preg_match_all($this->getTagRegex(), $tree->getText(), $matches, PREG_PATTERN_ORDER);
            // check region start e.g. %REGION:TEST:EXPR%
            if (false !== ($s = $this->findMatch($matches, $region.':'))) {
                $smatch = $matches[1][$s];
                $parts = explode(':', $smatch, 3);
                $id = $parts[1];
                $expr = $parts[2];
                // check region end e.g. %REGIONE:TEST%
                if (false !== ($e = $this->findMatch($matches, $region.'E:'.$id))) {
                    $ematch = $matches[1][$e];
                    // extract region
                    $region = $this->extract($tree, $bpos, $epos, $smatch, $ematch);
                    // region found
                    if (false !== $bpos && false !== $epos) {
                        // collect text between region mark
                        $rtext = null;
                        for ($i = $bpos; $i <= $epos; $i++) {
                            $rtext .= $tree->getMainGroup()->getChildAt($i)->getNodeText();
                        }
                        // clean before start mark
                        $tag = $this->createTag($smatch);
                        if (false !== ($p = mb_strpos($rtext, $tag))) {
                            $rtext = mb_substr($rtext, $p);
                        }
                        // clean after end mark
                        $tag = $this->createTag($ematch);
                        if (false !== ($p = mb_strpos($rtext, $tag))) {
                            $rtext = mb_substr($rtext, 0, $p + mb_strlen($tag));
                        }
                        // replace source tree with new placeholder
                        $regionId = $this->createTag($this->createTag($region.':'.$id));
                        $tree->getMainGroup()->replaceTextEx($rtext, $regionId);
                        // add the result
                        $result[$regionId] = array('expr' => $expr, 'tree' => $region);

                        continue;
                    }
                }
            }
            break;
        }

        return $result;
    }

    /**
     * Build and parse template tree.
     *
     * @param \NTLAB\RtfTree\Node\Tree $tree  Template tree
     * @param mixed $objects  The objects
     * @param boolean $notify  Notify listener
     * @return \NTLAB\RtfTree\Node\Tree
     */
    protected function doBuild(Tree $tree, $objects, $notify = true)
    {
        $result = $this->createTree();

        $body = $this->extract($tree, $bpos, $epos);
        $eachs = $this->extractRegion($body, 'EACH');
        $tables = $this->extractRegion($body, 'TBL');
        // include header
        for ($i = 0; $i < $bpos; $i++) {
            $result->getMainGroup()->appendChild($tree->getMainGroup()->getChildAt($i)->cloneNode());
        }
        // process body
        $this->getScript()
            ->setObjects($objects)
            ->each(function(Script $script, RtfFiler $_this) use ($result, $body, $eachs, $tables) {
                // replace regular tags
                $_this->appendTree($result, $_this->replaceTags($body));
                // replace each tags
                // replace tables tags
            }, $notify)
        ;
        // include footer
        for ($i = $epos + 1; $i < count($tree->getMainGroup()->getChildren()); $i++) {
            $result->getMainGroup()->appendChild($tree->getMainGroup()->getChildAt($i)->cloneNode());
        }

        return $result;
    }

    /**
     * Append tree nodes to another tree.
     * 
     * @param \NTLAB\RtfTree\Node\Tree $dest  Destination tree
     * @param \NTLAB\RtfTree\Node\Tree $source  Source tree
     * @return \NTLAB\Report\Util\RtfFiler
     */
    public function appendTree(Tree $dest, Tree $source)
    {
        foreach ($source->getMainGroup()->getChildren() as $node) {
            $dest->getMainGroup()->appendChild($node->cloneNode());
        }

        return $this;
    }

    /**
     * Replace all tags in tree.
     *
     * @param \NTLAB\RtfTree\Node\Tree $tree  The template body tree
     * @return \NTLAB\RtfTree\Node\Tree
     */
    public function replaceTags(Tree $tree)
    {
        $caches = array();
        $result = $tree->cloneTree();
        preg_match_all($this->getTagRegex(), $result->getText(), $matches, PREG_PATTERN_ORDER);
        for ($i = 0; $i < count($matches[0]); $i++) {
            $match = $matches[0][$i];
            $tag = $matches[1][$i];
            if (isset($caches[$tag])) {
                $value = $caches[$tag];
            } else {
                if (!$this->getScript()->getVar($value, $tag, $this->script->getContext())) {
                    $value = $this->getScript()->evaluate($tag);
                }
                $caches[$tag] = $value;
            }
            $result->getMainGroup()->replaceTextEx($match, $value);
        }

        return $result;
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