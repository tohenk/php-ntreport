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
use NTLAB\Script\Core\Script;

class Csv extends Report
{
    /**
     * @var array
     */
    protected $columns = array();

    /**
     * @var string
     */
    protected $delimeter = ';';

    /**
     * @var boolean
     */
    protected $withHeader = true;

    /**
     * @var resource
     */
    protected $handle = null;

    /**
     * @var boolean
     */
    protected $header = null;

    protected function initialize()
    {
        $this->hasTemplate = false;
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
        $this->template = $this->nodeAttr($node, 'name');
        $this->delimeter = $this->nodeAttr($node, 'delimeter', $this->delimeter);
        $this->withHeader = (bool) $this->nodeAttr($node, 'header', $this->withHeader);
    }

    protected function configureColumns(\DOMNode $node)
    {
        foreach ($node->childNodes as $column) {
            if ($name = $this->nodeAttr($column, 'name')) {
                $this->columns[$name] = $this->nodeAttr($column, 'value');
            }
        }
    }

    protected function build()
    {
        $this->handle = fopen('php://memory', 'r+');
        $this->header = false;
        $this->getScript()
            ->setObjects($this->result)
            ->each(function(Script $script, Csv $_this) {
                $_this->doBuild();
            })
        ;

        return $this->handle;
    }

    public function doBuild()
    {
        if (is_resource($this->handle)) {
            $header = array();
            $data = array();
            foreach ($this->columns as $label => $expr) {
                if (!$this->header) {
                    $header[] = $label;
                }
                $data[] = $this->getScript()->evaluate($expr);
            }
            if ($this->withHeader && !$this->header) {
                $this->header = true;
                fputcsv($this->handle, $header, $this->delimeter);
            }
            fputcsv($this->handle, $data, $this->delimeter);
        }
    }
}