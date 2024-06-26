<?php

/*
 * The MIT License
 *
 * Copyright (c) 2014-2024 Toha <tohenk@yahoo.com>
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

namespace NTLAB\Report;

class Template implements \ArrayAccess
{
    /**
     * @var \NTLAB\Report\Report
     */
    protected $report;

    /**
     * @var string
     */
    protected $content = null;

    /**
     * @var array
     */
    protected $properties = [];

    /**
     * Constructor.
     *
     * @param \NTLAB\Report\Report $report
     */
    public function __construct($report)
    {
        $this->report = $report;
    }

    /**
     * Set property.
     *
     * @param string $name
     * @param mixed $value
     * @return \NTLAB\Report\Template
     */
    public function setProperty($name, $value)
    {
        $this->properties[$name] = $value;
        return $this;
    }

    /**
     * Get property.
     *
     * @param string $name
     * @param mixed $default
     * @return mixed
     */
    public function getProperty($name, $default = null)
    {
        return isset($this->properties[$name]) ? $this->properties[$name] : $default;
    }

    /**
     * Get owner.
     *
     * @return \NTLAB\Report\Report
     */
    public function getReport()
    {
        return $this->report;
    }

    /**
     * Set template content.
     *
     * @param string $content
     * @return \NTLAB\Report\Template
     */
    public function setContent($content)
    {
        $this->content = $content;
        return $this;
    }

    /**
     * Get template content.
     *
     * @return string
     */
    public function getContent()
    {
        if (is_resource($this->content)) {
            // get resource size
            fseek($this->content, 0, SEEK_END);
            $size = ftell($this->content);
            // rewind to first position
            fseek($this->content, 0);
            return fread($this->content, $size);
        }
        if (is_callable($this->content)) {
            return call_user_func($this->content, $this);
        }
        return $this->content;
    }

    public function offsetExists(mixed $offset): bool
    {
        return isset($this->properties[$offset]);
    }

    public function offsetGet(mixed $offset): mixed
    {
        return $this->properties[$offset];
    }

    public function offsetSet(mixed $offset, mixed $value): void
    {
        $this->properties[$offset] = $value;
    }

    public function offsetUnset(mixed $offset): void
    {
        unset($this->properties[$offset]);
    }

    public function __toString()
    {
        return $this->getContent();
    }
}
