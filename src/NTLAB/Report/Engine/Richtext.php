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

namespace NTLAB\Report\Engine;

use NTLAB\Report\Report;
use NTLAB\Report\Util\Rtf\FilerInterface;
use NTLAB\Report\Util\Rtf\FilerLegacy;
use NTLAB\Report\Util\Rtf\FilerTree;

class Richtext extends Report
{
    /**
     * @var \NTLAB\Report\Util\Rtf\FilerInterface
     */
    protected $filer = null;

    /**
     * @var boolean
     */
    protected $single = null;

    /**
     * @var string
     */
    protected static $defaultFiler = null;

    /**
     * Set rich text default filer.
     *
     * @param string $class  Filer class name
     */
    public static function setDefaultFiler($class)
    {
        static::$defaultFiler = $class;
    }

    protected function configure(\DOMNodeList $nodes)
    {
        foreach ($nodes as $node) {
            switch (strtolower($node->nodeName)) {
                case 'template':
                    $this->configureParams($node);
                    break;

                case 'columns':
                    $this->configureColumns($node);
                    break;
            }
        }
    }

    protected function configureParams(\DOMNode $node)
    {
        $this->template = $this->nodeAttr($node, 'file');
        $this->single = $this->nodeAttr($node, 'single');
    }

    protected function configureColumns(\DOMNode $node)
    {
        // do nothing
    }

    /**
     * Get rich text filer.
     *
     * @return \NTLAB\Report\Util\Rtf\FilerInterface
     */
    protected function getFiler()
    {
        if (null === $this->filer) {
            if (null === static::$defaultFiler) {
                static::$defaultFiler = FilerLegacy::class;
            }
            $this->filer = new static::$defaultFiler();
            $this->filer->setScript($this->getScript());
        }

        return $this->filer;
    }

    protected function build()
    {
        $objects = $this->result;
        // is build for single content?
        if ($objects && $this->single) {
            $objects = array($objects[0]);
        }

        return $this->getFiler()->build($this->templateContent, $objects);
    }
}