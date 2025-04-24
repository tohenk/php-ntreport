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
use NTLAB\Report\Filer\Document as DocumentFiler;
use NTLAB\Report\Session\Session;

class Word extends Report
{
    public const ID = 'doc';

    /**
     * @var boolean
     */
    protected $single = null;

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
        $this->single = $this->nodeAttr($node, 'single');
    }

    protected function configureColumns(\DOMNode $node)
    {
        // do nothing
    }

    protected function build()
    {
        $content = null;
        $error = null;
        $this->session->load();
        $template = $this->templateContent->getContent();
        if (!$tmpl = $this->session->read(Session::TEMPLATE)) {
            if ($template) {
                $tmpl = $this->session->getTempFile($this->template);
                file_put_contents($tmpl, $template);
                $this->session->store(Session::TEMPLATE, $tmpl);
            } else {
                $this->status = static::STATUS_ERR_TMPL;
            }
        }
        if (null === $this->status) {
            try {
                DocumentFiler::setTempDir($this->session->createWorkDir());
                $objects = $this->result;
                // is build for single content?
                if ($objects && $this->single) {
                    $objects = [$objects[0]];
                }
                $filer = new DocumentFiler($tmpl);
                $content = $filer
                    ->setSession($this->session)
                    ->build(null, $objects);
                unset($filer);
            } catch (\Exception $e) {
                $error = $e;
            }
            $this->session->clean();
            if (null !== $error) {
                throw $error;
            }
        }

        return $content;
    }
}
