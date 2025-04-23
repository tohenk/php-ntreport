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

use NTLAB\Report\Report;
use NTLAB\Report\Filer\Spreadsheet as SpreadsheetFiler;

class Excel extends Report
{
    public const ID = 'excel';

    /**
     * @var \NTLAB\Report\Filer\Spreadsheet
     */
    protected $filer = null;

    protected function configure(\DOMNodeList $nodes)
    {
        $this->filer = new SpreadsheetFiler();
        $this->filer->setScript($this->getScript());
        foreach ($nodes as $node) {
            switch (strtolower($node->nodeName)) {
                case 'template':
                    $this->configureTemplate($node);
                    break;
                case 'bands':
                    $this->configureBands($node);
                    break;
                case 'datas':
                    $this->configureDatas($node);
                    break;
            }
        }
    }

    protected function configureTemplate(\DOMNode $node)
    {
        $this->template = $this->nodeAttr($node, 'name');
        $this->filer->setSheet($this->nodeAttr($node, 'sheet'));
        $this->filer->setFieldSign($this->nodeAttr($node, 'field-sign', $this->filer->getFieldSign()));
        $this->filer->setAutoFit($this->nodeAttr($node, 'autofit', $this->filer->getAutoFit()));
        $this->filer->setRowHeight($this->nodeAttr($node, 'rowheight', $this->filer->getRowHeight()));
    }

    protected function configureBands(\DOMNode $node)
    {
        foreach ($node->childNodes as $band) {
            $this->filer->addBand($this->nodeAttr($band, 'type'), $this->nodeAttr($band, 'range'));
        }
    }

    protected function configureDatas(\DOMNode $node)
    {
        foreach ($node->childNodes as $data) {
            $this->filer->addData($this->nodeAttr($data, 'name'), $this->nodeAttr($data, 'value'));
        }
    }

    protected function getTempFile()
    {
        $ext = substr($this->template, strrpos($this->template, '.'));
        $name = substr(sha1(rand(1, 99999)), 0, 8);

        return sys_get_temp_dir().DIRECTORY_SEPARATOR.$name.$ext;
    }

    protected function build()
    {
        $content = null;
        $error = null;
        if ($template = $this->templateContent->getContent()) {
            $tempfile = $this->getTempFile();
            try {
                file_put_contents($tempfile, $template);
                $content = $this->filer->build($tempfile, $this->result);
            } catch (\Exception $e) {
                $error = $e;
            }
            unlink($tempfile);
            if (null !== $error) {
                throw $error;
            }
        } else {
            $this->status = static::STATUS_ERR_TMPL;
        }

        return $content;
    }
}
