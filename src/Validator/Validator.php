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

namespace NTLAB\Report\Validator;

use NTLAB\Report\Report;
use NTLAB\Script\Core\Script;

class Validator
{
    /**
     * @var string
     */
    protected $name = null;

    /**
     * @var string
     */
    protected $expr = null;

    /**
     * @var \NTLAB\Report\Report
     */
    protected $report = null;

    /**
     * Constructor.
     *
     * @param string $name  The validator name
     * @param string $expr  The expression
     */
    public function __construct($name, $expr)
    {
        $this->name = $name;
        $this->expr = $expr;
    }

    /**
     * Set the report object.
     *
     * @param \NTLAB\Report\Report $report  The report object
     * @return \NTLAB\Report\Validator\Validator
     */
    public function setReport(Report $report)
    {
        $this->report = $report;

        return $this;
    }

    /**
     * Check the validity.
     *
     * @return bool
     */
    public function isValid()
    {
        $this->report->getScript()->setContext($this->report->getObject());
        $value = $this->report->getScript()->evaluate($this->expr);

        return Script::asBool($value);
    }
}