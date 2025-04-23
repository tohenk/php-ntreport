<?php

/*
 * The MIT License
 *
 * Copyright (c) 2014-2025 Toha <tohenk@yahoo.com>
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

class Table extends Extractor
{
    public const TROWD = 'trowd';
    public const ROW = 'row';

    /**
     * Extract table region.
     *
     * @param \NTLAB\RtfTree\Node\Tree $tree  Input tree
     * @return \NTLAB\Report\Util\Rtf\Extractor\Paragraph
     */
    public static function getTable(Tree $tree)
    {
        $extractor = new self();
        $extractor->regions = $extractor->extractRegion($tree, 'TBL');

        return $extractor;
    }

    /**
     * Constructor.
     */
    public function __construct()
    {
        $this->beginKey = static::TROWD;
        $this->endKey = static::ROW;
    }

    /**
     * (non-PHPdoc)
     * @see \NTLAB\Report\Util\Rtf\Extractor\Extractor::isMatchEndKey()
     */
    protected function isMatchEndKey(Node $node)
    {
        if ($node->is(Node::GROUP)) {
            if ($node = $node->selectSingleChildNode($this->endKey)) {
                return true;
            }
        }

        return false;
    }

    /**
     * (non-PHPdoc)
     * @see \NTLAB\Report\Util\Rtf\Extractor\Extractor::ensureEndKey()
     */
    protected function ensureEndKey(Tree $tree, &$position)
    {
        parent::ensureEndKey($tree, $position);
        if ($node = $tree->getMainGroup()->getChildAt($position)) {
            if ($this->isMatchEndKey($node) &&
                ($nextNode = $tree->getMainGroup()->getChildAt($position + 1)) &&
                $nextNode->isEquals(Paragraph::PARAGRAPH)) {
                $position++;
            }
        }

        return $this;
    }
}
