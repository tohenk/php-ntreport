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

namespace NTLAB\Report\Util\Rtf\Extractor;

use NTLAB\RtfTree\Node\Tree;
use NTLAB\RtfTree\Node\Node;

class Extractor
{
    const TAG_SIGN = '%';

    /**
     * @var string
     */
    protected $beginMark;

    /**
     * @var string
     */
    protected $endMark;

    /**
     * @var int
     */
    protected $beginPos;

    /**
     * @var int
     */
    protected $endPos;

    /**
     * @var string
     */
    protected $beginKey;

    /**
     * @var string
     */
    protected $endKey;

    /**
     * @var \NTLAB\RtfTree\Node\Tree
     */
    protected $result;

    /**
     * @var array
     */
    protected $regions = array();

    /**
     * @var string
     */
    protected static $re;

    /**
     * Get the enclosing tag signs.
     * The result is an array which contains
     * the opening and closing tag sign array(opening, closing).
     *
     * @return array
     */
    public static function getTags()
    {
        $stag = substr(static::TAG_SIGN, 0, 1);
        $etag = strlen(static::TAG_SIGN) > 1 ? substr(static::TAG_SIGN, 1, 1) : $stag;

        return array($stag, $etag);
    }

    /**
     * Get tag regular expression pattern.
     *
     * @return string
     */
    public static function getTagRegex()
    {
        if (null == static::$re) {
            $tags = static::getTags();
            static::$re = sprintf('/%1$s([^%1$s]+)%2$s/', $tags[0], $tags[1]);
        }

        return static::$re;
    }

    /**
     * Create a tag.
     *
     * @param string $tag  The tag
     * @return string
     */
    public static function createTag($tag)
    {
        $tags = static::getTags();

        return $tags[0].$tag.$tags[1];
    }

    /**
     * Find tag in the tree.
     *
     * @param \NTLAB\RtfTree\Node\Tree $tree  The tree 
     * @param string $tag  Tag to find
     * @param int $size  The nodes size found for matched tag
     * @param int $start  Start position
     * @return int
     */
    public static function findTag(Tree $tree, $tag, &$size, $start = null)
    {
        $tag = static::createTag($tag);
        if (!is_int($start)) {
            $start = 0;
        }
        $text = null;
        for ($i = $start; $i < count($tree->getMainGroup()->getChildren()); $i++) {
            if (!($node = $tree->getMainGroup()->getChildAt($i))) {
                continue;
            }
            $plain = $node->getPlainText();
            // tag found in single node
            if (false !== mb_strpos($plain, $tag)) {
                $size = 1;
                return $i;
            }
            // check for combined text
            $text .= $plain;
            if (false !== mb_strpos($text, $tag)) {
                $text = $plain;
                $size = 1;
                while (true) {
                    $i--;
                    if ($i <= $start) {
                        break;
                    }
                    $size++;
                    $text = $tree->getMainGroup()->getChildAt($i)->getPlainText().$text;
                    if (false !== mb_strpos($text, $tag)) {
                        return $i;
                    }
                }
            }
        }

        return false;
    }

    /**
     * Create rtf tree.
     *
     * @return \NTLAB\RtfTree\Node\Tree
     */
    public static function createTree()
    {
        $tree = new Tree();
        $tree->addMainGroup();

        return $tree;
    }

    /**
     * Copy tree nodes.
     *
     * @param \NTLAB\RtfTree\Node\Tree $dest  Destination tree
     * @param \NTLAB\RtfTree\Node\Tree $source  Source tree
     * @param int $start  Start index
     * @param int $end  End index
     * @return \NTLAB\Report\Util\Rtf\Extractor\Extractor
     */
    public static function copyTree(Tree $dest, Tree $source, $start, $end)
    {
        for ($i = $start; $i <= $end; $i++) {
            $dest->getMainGroup()->appendChild($source->getMainGroup()->getChildAt($i)->cloneNode());
        }
    }

    /**
     * Append tree nodes to another tree.
     * 
     * @param \NTLAB\RtfTree\Node\Tree $dest  Destination tree
     * @param \NTLAB\RtfTree\Node\Tree $source  Source tree
     */
    public static function appendTree(Tree $dest, Tree $source)
    {
        foreach ($source->getMainGroup()->getChildren() as $node) {
            $dest->getMainGroup()->appendChild($node->cloneNode());
        }
    }

    /**
     * Append tree nodes to another tree by inserting at specified
     * position.
     * 
     * @param \NTLAB\RtfTree\Node\Tree $dest  Destination tree
     * @param \NTLAB\RtfTree\Node\Tree $source  Source tree
     * @param int $position  Start position
     */
    public static function insertTree(Tree $dest, Tree $source, $position)
    {
        foreach ($source->getMainGroup()->getChildren() as $node) {
            $dest->getMainGroup()->insertChild($position, $node);
            $position++;
        }
    }

    /**
     * Get matched begin position index.
     *
     * @return int
     */
    public function getBeginPos()
    {
        return $this->beginPos;
    }

    /**
     * Get matched end position index.
     *
     * @return int
     */
    public function getEndPos()
    {
        return $this->endPos;
    }

    /**
     * Get result tree.
     *
     * @return \NTLAB\RtfTree\Node\Tree
     */
    public function getResult()
    {
        return $this->result;
    }

    /**
     * Get result regions.
     * 
     * @return array
     */
    public function getRegions()
    {
        return $this->regions;
    }

    /**
     * Extract template for begin and end mark. If begin or ending mark is not
     * found, return all paragraphs within template.
     *
     * @param \NTLAB\RtfTree\Node\Tree $tree  Template tree
     * @param int $this->beginPos  Begin mark position found
     * @param int $this->endPos  End mark position found
     * @param string $bmark  Begining mark
     * @param string $emark  Ending mark
     * @return \NTLAB\RtfTree\Node\Tree
     * @throws \InvalidArgumentException
     */
    public function extract(Tree $tree)
    {
        $start = $this->getStartIndex($tree);
        $result = $this->createTree();

        $ssize = null;
        $esize = null;
        $this->beginPos = $this->findTag($tree, $this->beginMark, $ssize, $start);
        $this->endPos = $this->findTag($tree, $this->endMark, $esize, $this->beginPos);
        // begin and end mark found
        if (false !== $this->beginPos && false !== $this->endPos) {
            // ensure first match contain begin key
            if ($this->beginKey) {
                $this->ensureBeginKey($tree, $start, $this->beginPos);
            }
            $this->endPos += $esize - 1;
            // ensure last match contain end key
            if ($this->endKey) {
                $this->ensureEndKey($tree, $this->endPos);
            }
            // copy marked
            $this->copyTree($result, $tree, $this->beginPos, $this->endPos);
            // remove begin and end mark
            $result->getMainGroup()->replaceTextEx($this->createTag($this->beginMark), '');
            $result->getMainGroup()->replaceTextEx($this->createTag($this->endMark), '');
        } else {
            $this->beginPos = $start;
            $this->endPos = count($tree->getMainGroup()->getChildren()) - 1;
            // copy all
            $this->copyTree($result, $tree, $this->beginPos, $this->endPos);
        }

        return $result;
    }

    /**
     * Extract region from tree.
     *
     * @param \NTLAB\RtfTree\Node\Tree $tree  The tree to extract
     * @param string $region  Region name
     * @return array
     */
    public function extractRegion(Tree $tree, $region)
    {
        $result = array();
        while (true) {
            $matches = null;
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
                    $this->beginMark = $smatch;
                    $this->endMark = $ematch;
                    $rtree = $this->extract($tree);
                    // region found
                    if (false !== $this->beginPos && false !== $this->endPos) {
                        $regionId = $region.'_'.$id;
                        // remove matched range
                        $tree->getMainGroup()->getChildren()->removeAt($this->beginPos, $this->endPos - $this->beginPos + 1);
                        // insert placeholder
                        $tree->getMainGroup()->insertChild($this->beginPos, Node::create(Node::TEXT, $this->createTag($regionId)));
                        // add the result
                        $result[$regionId] = array($expr, $rtree);

                        continue;
                    }
                }
            }
            break;
        }

        return $result;
    }

    /**
     * Get start index for extraction.
     *
     * @param \NTLAB\RtfTree\Node\Tree $tree  Input tree
     * @return int
     */
    protected function getStartIndex(Tree $tree)
    {
        return 0;
    }

    /**
     * Ensure start index contain designed key.
     *
     * @param \NTLAB\RtfTree\Node\Tree $tree  Input tree
     * @param int $start  Start index
     * @param int $position  Current position index
     * @return \NTLAB\Report\Util\Rtf\Extractor\Extractor
     */
    protected function ensureBeginKey(Tree $tree, $start, &$position)
    {
        while (true) {
            if ($position === $start) {
                break;
            }
            $node = $tree->getMainGroup()->getChildAt($position);
            if ($this->isMatchBeginKey($node)) {
                break;
            }
            $position--;
        }

        return $this;
    }

    /**
     * Check if node matched begin key.
     *
     * @param \NTLAB\RtfTree\Node\Node $node  The node
     * @return boolean
     */
    protected function isMatchBeginKey(Node $node)
    {
        return $node->isEquals($this->beginKey);
    }

    /**
     * Ensure end index contain designed key.
     *
     * @param \NTLAB\RtfTree\Node\Tree $tree  Input tree
     * @param int $position  Current position index
     * @return \NTLAB\Report\Util\Rtf\Extractor\Extractor
     */
    protected function ensureEndKey(Tree $tree, &$position)
    {
        $last = count($tree->getMainGroup()->getChildren()) - 1;
        while (true) {
            if ($position === $last) {
                break;
            }
            $node = $tree->getMainGroup()->getChildAt($position);
            if ($this->isMatchEndKey($node)) {
                break;
            }
            $position++;
        }

        return $this;
    }

    /**
     * Check if node matched end key.
     *
     * @param \NTLAB\RtfTree\Node\Node $node  The node
     * @return boolean
     */
    protected function isMatchEndKey(Node $node)
    {
        return $node->isEquals($this->endKey);
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
            $tag = $matches[1][$i];
            if (0 === mb_strpos($tag, $prefix)) {
                return $i;
            }
        }

        return false;
    }
}