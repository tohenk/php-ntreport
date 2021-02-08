<?php

/*
 * The MIT License
 *
 * Copyright (c) 2014-2021 Toha <tohenk@yahoo.com>
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

namespace NTLAB\Report\Filer;

/**
 * Quick and dirty filer for RTF using regular expression.
 *
 * @author Toha
 */
class Rtf extends RtfTag
{
    /**
     * @var string
     */
    protected $header = null;

    /**
     * @var string
     */
    protected $footer = null;

    /**
     * Constructor.
     */
    public function __construct()
    {
        $this->break = '\page';
        $this->notifyContextChange = true;
    }

    /**
     * Extract a template into parts: header, body, and footer.
     *
     * @param string $template  The template
     * @param string $sregex  The begin tag regex
     * @param string $eregex  The end tag regex
     * @return bool|array If success, return array(header, body, footer), false otherwise
     */
    protected function extract($template, $sregex = 'BEGIN', $eregex = 'END')
    {
        if (is_resource($template)) {
            $template = stream_get_contents($template);
        }
        $smatch = null;
        $ematch = null;
        $s = $this->findTag($template, $sregex, $smatch);
        $e = $this->findTag($template, $eregex, $ematch);
        if (is_array($parts = $this->splitParts($template, $s, $e, $smatch, $ematch))) {
            return $parts;
        }
        return false;
    }

    /**
     * {@inheritDoc}
     * @see \NTLAB\Report\Filer\RtfTag::beginDoc()
     */
    protected function beginDoc()
    {
        $this->content = $this->header;
        return $this;
    }

    /**
     * {@inheritDoc}
     * @see \NTLAB\Report\Filer\RtfTag::endDoc()
     */
    protected function endDoc()
    {
        $this->content .= $this->footer;
        return $this;
    }

    /**
     * {@inheritDoc}
     * @see \NTLAB\Report\Filer\RtfTag::build()
     */
    public function build($template, $objects)
    {
        if (is_array($parts = $this->extract($template))) {
            $this->header = $parts[0];
            $this->footer = $parts[2];
            return parent::build($parts[1], $objects);
        }
    }
}