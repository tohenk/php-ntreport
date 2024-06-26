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

namespace NTLAB\Report\Script;

use NTLAB\Script\Provider\ProviderInterface;

class ProviderReport implements ProviderInterface
{
    /**
     * @var \NTLAB\Script\Core\Module[]
     */
    protected $modules = [];

    /**
     * @var \NTLAB\Report\Script\ProviderReport
     */
    protected static $instance = null;

    /**
     * Get instance.
     *
     * @return \NTLAB\Report\Script\ProviderReport
     */
    public static function getInstance()
    {
        if (null === static::$instance) {
            static::$instance = new self();
        }
        return static::$instance;
    }

    /**
     * Constructor.
     */
    public function __construct()
    {
        $this->modules = [
            new ReportCore(),
            new ReportTag(),
            new ReportExcel(),
        ];
    }

    /**
     * (non-PHPdoc)
     * @see \NTLAB\Script\Provider\ProviderInterface::getModules()
     */
    public function getModules()
    {
        return $this->modules;
    }
}