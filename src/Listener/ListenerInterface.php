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

namespace NTLAB\Report\Listener;

use NTLAB\Report\Data\Data;
use NTLAB\Report\Parameter\Parameter;

interface ListenerInterface
{
    /**
     * Format parameter title.
     *
     * @param \NTLAB\Report\Parameter\Parameter $parameter  Parameter title to format
     * @param string $title  Resulting title
     * @return boolean
     */
    public function formatTitle(Parameter $parameter, &$title);

    /**
     * Apply parameter to underlying data handler.
     *
     * @param \NTLAB\Report\Parameter\Parameter $parameter  The parameter to apply
     * @param NTLAB\Report\Data\Data $data  Data handler
     * @param array $result  The result array
     * @return boolean
     */
    public function applyParameter(Parameter $parameter, Data $data, &$result);
}
