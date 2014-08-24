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

namespace NTLAB\Report\Util\Extractor;

use NTLAB\RtfTree\Node\Tree;

class Paragraph extends Extractor
{
    const PARAGRAPH = 'pard';

    /**
     * Extract body.
     *
     * @param \NTLAB\RtfTree\Node\Tree $tree  Input tree
     * @return \NTLAB\Report\Util\Extractor\Paragraph
     */
    public static function getBody(Tree $tree)
    {
        $extractor = new self();
        $extractor->result = $extractor->extract($tree);

        return $extractor;
    }

    /**
     * Extract region.
     *
     * @param \NTLAB\RtfTree\Node\Tree $tree  Input tree
     * @param string $region  Region name
     * @return \NTLAB\Report\Util\Extractor\Paragraph
     */
    public static function getRegion(Tree $tree, $region)
    {
        $extractor = new self();
        $extractor->regions = $extractor->extractRegion($tree, $region);

        return $extractor;
    }

    /**
     * Constructor.
     *
     * @param string $beginMark  Begin mark
     * @param string $endMark  End mark
     */
    public function __construct($beginMark = 'BEGIN', $endMark = 'END')
    {
        $this->beginKey = static::PARAGRAPH;
        $this->beginMark = $beginMark;
        $this->endMark = $endMark;
    }

    /**
     * (non-PHPdoc)
     * @see \NTLAB\Report\Util\Extractor\Extractor::getStartIndex()
     */
    protected function getStartIndex(Tree $tree)
    {
        if (!($node = $tree->getMainGroup()->selectSingleNode(static::PARAGRAPH))) {
            throw new \InvalidArgumentException('No paragraph found.');
        }

        return $node->getNodeIndex();
    }
}