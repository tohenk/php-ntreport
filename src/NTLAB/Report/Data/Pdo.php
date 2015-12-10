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

use NTLAB\Report\Parameter\Parameter;
use NTLAB\Report\Parameter\Date;
use NTLAB\Report\Query\Pdo as PdoQuery;

class Pdo extends Data
{
    /**
     * @var array
     */
    protected $conds = array();

    /**
     * @var array
     */
    protected $orders = array();

    /**
     * @var array
     */
    protected $params = array();

    /**
     * @var boolean
     */
    protected static $supported = null;

    /**
     * Check if propel pdo data is supported.
     *
     * @return boolean
     */
    public static function isSupported()
    {
        if (null == static::$supported) {
            static::$supported = extension_loaded('pdo');
        }

        return static::$supported;
    }

    /**
     * (non-PHPdoc)
     * @see \NTLAB\Report\Data\Data::canHandle()
     */
    public function canHandle($source)
    {
        return $this->isSupported() && false !== stripos($source, 'SELECT');
    }

    /**
     * (non-PHPdoc)
     * @see \NTLAB\Report\Data\Data::addCondition()
     */
    public function addCondition(Parameter $parameter)
    {
        foreach ($this->convertParameter($parameter) as $params) {
            list($column, $operator, $value) = $params;
            if (null === $value && null === $operator) {
                $this->conds[] = $column;
            } else {
                $pname = sprintf(':p%d', count($this->params) + 1);
                $this->conds[] = sprintf('%s %s %s', $column, $operator, $pname);
                $this->params[$pname] = $value;
            }
        }
    }

    /**
     * (non-PHPdoc)
     * @see \NTLAB\Report\Data\Data::addOrder()
     */
    public function addOrder($column, $direction, $format = null)
    {
        $this->orders[] = trim($this->applyFormat($format, $column).' '.$direction);
    }

    /**
     * (non-PHPdoc)
     * @see \NTLAB\Report\Data\Data::fetch()
     */
    public function fetch()
    {
        $sql = strtr($this->getSource(), array(
            '%COND%' => count($this->conds) ? implode(' AND ', $this->conds) : '1',
            '%ORDER%' => count($this->orders) ? 'ORDER BY '.implode(', ', $this->orders) : ''
        ));
        $sql = $this->report->getScript()->evaluate($sql);

        return PdoQuery::getInstance()->query($sql, $this->params);
    }

    /**
     * Format a date as PDO statement query.
     *
     * @param string $column  Column name
     * @param string $dateType  Date type
     * @param int $dateValue  Date timestamp value
     * @param string $operator  Condition operator
     * @return string
     */
    protected function formatDate($column, $dateType, $dateValue, $operator = null)
    {
        if (Date::DATE === $dateType) {
            $dateValue = date('Y-m-d', $dateValue);
        } else if (Date::MONTH === $dateType) {
            $dateValue = date('Y-m', $dateValue);
        } else if (Date::YEAR === $dateType) {
            $dateValue = date('Y', $dateValue);
        }

        return sprintf('%1$s %2$s %4$s%3$s%4$s',
            sprintf('SUBSTRING(%s, %d, %d)', $column, 1, strlen($dateValue)),
            $operator ? $operator : '=',
            $dateValue,
            '"'
        );
    }

    /**
     * (non-PHPdoc)
     * @see \NTLAB\Report\Data\Data::doConvert()
     */
    protected function doConvert(Parameter $parameter, &$result)
    {
        $column = $parameter->getColumn();
        $operator = $parameter->getOperator();
        $value = $parameter->getCurrentValue();

        // convert parameter to SQL
        switch ($parameter->getType())
        {
            case 'static':
                if ($value === 'NULL') {
                    $column = sprintf('%s IS NULL', $column);
                    $operator = null;
                    $value = null;
                } else if ($value === 'NOT_NULL') {
                    $column = sprintf('%s IS NOT NULL', $column);
                    $operator = null;
                    $value = null;
                }
                break;

            case 'date':
            case 'dateonly':
                $column = $this->formatDate($parameter->getRealColumn(), $parameter->getDateTypeValue(), $value, $operator);
                $operator = null;
                $value = null;
                break;

            case 'daterange':
                $column = sprintf('(%s AND %s)',
                    $this->formatDate($parameter->getRealColumn(), $parameter->getDateTypeValue(), $value, '>='),
                    $this->formatDate($parameter->getRealColumn(), $parameter->getDateTypeValue(), $parameter->getCurrentValue2(), '<=')
                );
                $operator = null;
                $value = null;
                break;
        }
        $result[] = $this->createParam($column, $operator, $value);
    }
}