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

namespace NTLAB\Report\Parameter;

use NTLAB\Script\Core\Manager;
use NTLAB\Script\Context\ArrayVar;
use Propel\Runtime\DataFetcher\DataFetcherInterface;

class Reference extends Parameter
{
    const ID = 'ref';

    /**
     * Get reference value.
     *
     * @param string $value  The reference value
     * @return string
     */
    public function getText($value)
    {
        $values = $this->getValues();
        if (array_key_exists($value, $values)) {
            return $values[$value];
        }
    }

    public function getValue()
    {
        if (null == ($value = parent::getValue()) && count($values = $this->getValues())) {
            $keys = array_keys($values);
            $value = array_shift($keys);
        }

        return $value;
    }

    protected function getDefaultTitle()
    {
        return $this->getText($this->getValue());
    }
}