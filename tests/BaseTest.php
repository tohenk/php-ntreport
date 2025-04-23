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

namespace NTLAB\Report\Test;

use PHPUnit\Framework\TestCase;

abstract class BaseTest extends TestCase
{
    protected function getFixtureDir()
    {
        return __DIR__.'/fixtures/';
    }

    protected function getResultDir()
    {
        return __DIR__.'/result/';
    }

    protected function getOutDir()
    {
        return __DIR__.'/out/';
    }

    protected function loadFixture($name)
    {
        return file_get_contents($this->getFixtureDir().$name);
    }

    protected function loadResult($name)
    {
        return file_get_contents($this->getResultDir().$name);
    }

    protected function saveOut($content, $filename)
    {
        file_put_contents($this->getOutDir().$filename, $content);
    }
}
