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

namespace NTLAB\Report\Form;

use NTLAB\Report\Report;

abstract class FormBuilder
{
    /**
     * @var \NTLAB\Report\Report
     */
    protected $report;

    /**
     * Get owner report.
     *
     * @return \NTLAB\Report\Report
     */
    public function getReport()
    {
        return $this->report;
    }

    /**
     * Set owner report.
     *
     * @param \NTLAB\Report\Report $report  Owner report
     * @return \NTLAB\Report\Form\FormBuilder
     */
    public function setReport(Report $report)
    {
        $this->report = $report;

        return $this;
    }

    /**
     * Build report form.
     *
     * @param array $defaults  Default values
     * @return \NTLAB\Report\Form\FormInterface
     */
    abstract public function build($defaults = []);

    /**
     * Normalize and remove illegal characters.
     *
     * @param string $name  Name to normalize
     * @return string
     */
    abstract public function normalize($name);
}
