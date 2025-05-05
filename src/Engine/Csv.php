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

namespace NTLAB\Report\Engine;

use NTLAB\Report\Filer\FilerInterface;
use NTLAB\Report\Report;
use NTLAB\Report\Session\Session;
use NTLAB\Script\Core\Script;

class Csv extends Report
{
    public const ID = 'csv';

    /**
     * @var array
     */
    protected $columns = [];

    /**
     * @var string
     */
    protected $delimiter = ';';

    /**
     * @var string
     */
    protected $enclosure = '"';

    /**
     * @var string
     */
    protected $escape = '\\';

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

    protected function doInitialize()
    {
        $this->acceptPartialObject = true;
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
        $this->delimiter = $this->nodeAttr($node, 'delimiter', $this->delimiter);
        $this->delimiter = $this->nodeAttr($node, 'delimeter', $this->delimiter);
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
        $content = null;
        $error = null;
        $this->session->load();
        try {
            $filename = $this->session->createWorkDir().DIRECTORY_SEPARATOR.$this->template;
            $this->handle = fopen($filename, 'a+');
            $this->header = false;
            $this->getScript()
                ->setObjects($this->result)
                ->each(function (Script $script, Csv $_this) {
                    $_this->doBuild();
                })
            ;
        } catch (\Exception $e) {
            $error = $e;
        }
        $iterator = $this->getScript()->getIterator();
        $done = $iterator->getRecNo() === $iterator->getRecCount();
        if ($done || $error) {
            $this->session->clean();
        }
        if (null !== $error) {
            throw $error;
        }
        if (!$done) {
            $this->session
                ->store(Session::POS, $iterator->getRecNo() - 1)
                ->save();
            $content = FilerInterface::PARTIAL;
        } else {
            $content = $this->handle;
        }

        return $content;
    }

    public function doBuild()
    {
        if (is_resource($this->handle)) {
            $iterator = $this->getScript()->getIterator();
            $first = $iterator->getStart() === 0;
            $header = [];
            $data = [];
            foreach ($this->columns as $label => $expr) {
                if ($first && !$this->header) {
                    $header[] = $label;
                }
                $data[] = $this->getScript()->evaluate($expr);
            }
            if ($this->withHeader && $first && !$this->header) {
                $this->header = true;
                fputcsv($this->handle, $header, $this->delimiter, $this->enclosure, $this->escape);
            }
            fputcsv($this->handle, $data, $this->delimiter, $this->enclosure, $this->escape);
        }
    }
}
