<?php

/*
 * The MIT License
 *
 * Copyright (c) 2016-2021 Toha <tohenk@yahoo.com>
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

namespace NTLAB\Report\Parameter;

class DateMonth extends Date
{
    const ID = 'month';

    /**
     * {@inheritDoc}
     * @see \NTLAB\Report\Parameter\Date::getDateTypeValue()
     */
    public function getDateTypeValue()
    {
        return static::MONTH;
    }

    public function getFieldName()
    {
        return parent::getFieldName().'_yr';
    }

    public function getFieldName2()
    {
        return parent::getFieldName().'_mo';
    }

    public function getValue()
    {
        if (null == ($year = $this->getFormValue($this->getFieldName())) && count($years = $this->getValues())) {
            $keys = array_keys($years);
            $year = array_shift($keys);
        }
        if (null == ($month = $this->getFormValue($this->getFieldName2()))) {
            $month = date('m');
        }
        // create datetime from original value
        if (is_numeric($year) && is_numeric($month)) {
            return mktime(0, 0, 0, (int) $month, 1, (int) $year);
        }
    }
}