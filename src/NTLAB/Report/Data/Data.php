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

namespace NTLAB\Report\Data;

use NTLAB\Report\Report;
use NTLAB\Report\Parameter\Parameter;

abstract class Data
{
    /**
     * @var array
     */
    protected static $handlers = array();

    /**
     * @var \NTLAB\Report\Report;
     */
    protected $report = null;

    /**
     * @var string
     */
    protected $source = null;

    /**
     * @var boolean
     */
    protected $distinct = false;

    /**
     * Register report data handler.
     *
     * @param string $class  The handler class
     */
    public static function register($class)
    {
        if (!in_array($class, self::$handlers)) {
            self::$handlers[] = $class;
        }
    }

    /**
     * Get report handler capable for report data.
     *
     * @param string $source  The report data
     * @return \NTLAB\Report\Data\Data
     */
    public static function getHandler($source)
    {
        foreach (self::$handlers as $handlerClass) {
            $handler = new $handlerClass();
            if ($handler->canHandle($source)) {
                $handler->source = $source;

                return $handler;
            }
            unset($handler);
        }

        return false;
    }

    /**
     * Is the class can handle source.
     *
     * @param string $source  The report data source
     * @return boolean
     */
    abstract public function canHandle($source);

    /**
     * Add condition to report data source.
     *
     * @param \NTLAB\Report\Parameter\Parameter $parameter  Report parameter
     */
    abstract public function addCondition(Parameter $parameter);

    /**
     * Add result ordering.
     *
     * @param string $column  The column name to order
     * @param string $direction  The order direction
     */
    abstract public function addOrder($column, $direction);

    /**
     * Fetch report data.
     *
     * @return array
     */
    abstract public function fetch();

    /**
     * Set the report object.
     *
     * @param \NTLAB\Report\Report $report  Report object
     */
    public function setReport($report)
    {
        $this->report = $report;
    }

    /**
     * Get report data source.
     *
     * @return string
     */
    public function getSource()
    {
        return $this->source;
    }

    /**
     * Is distinct.
     *
     * @return boolean
     */
    public function isDistinct()
    {
        return $this->distinct;
    }

    /**
     * Set the distinct value.
     *
     * @param boolean $value  The distinct value to set
     * @return \NTLAB\Report\Data\Data
     */
    public function setDistinct($value)
    {
        $this->distinct = (bool) $value;

        return $this;
    }

    /**
     * Get the real column name.
     *
     * @param string $column  The column name
     * @return string
     */
    public function getColumn($column)
    {
        return $column;
    }

    /**
     * Create parameter array.
     *
     * @param string $column  Column name
     * @param string $operator  The operator
     * @param mixed $value  The value
     * @return array
     */
    public function createParam($column, $operator, $value)
    {
        return array($column, $operator, $value);
    }

    /**
     * Call all listener for applying parameter.
     *
     * @param \NTLAB\Report\Parameter\Parameter $parameter  Parameter to apply
     * @param array $result  Result array
     * @return boolean
     */
    protected function applyParameter(Parameter $parameter, &$result)
    {
        foreach ($this->report->getListeners() as $listener) {
            if ($listener->applyParameter($parameter, $this, $result)) {
                return true;
            }
        }

        return false;
    }

    /**
     * Convert parameter to specific data format.
     *
     * @param \NTLAB\Report\Parameter\Parameter $parameter  The parameter to convert
     * @return array
     */
    protected function convertParameter(Parameter $parameter)
    {
        $result = array();
        if (!$this->applyParameter($parameter, $result)) {
            $this->doConvert($parameter, $result);
        }

        return $result;
    }

    /**
     * Internal convert parameter.
     *
     * @param \NTLAB\Report\Parameter\Parameter $parameter  The parameter to convert
     * @param array $result  Result array
     */
    abstract protected function doConvert(Parameter $parameter, &$result);
}