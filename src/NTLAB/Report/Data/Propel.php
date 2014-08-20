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

class Propel extends Data
{
    const COLUMN_CONCATENATOR = '.';

    /**
     * @var \ModelCriteria
     */
    protected $query = null;

    /**
     * @var \ModelCriteria[]
     */
    protected $items = array();

    /**
     * @var boolean
     */
    protected static $supported = null;

    /**
     * Check if propel data is supported.
     *
     * @return boolean
     */
    public static function isSupported()
    {
        if (null == static::$supported) {
            static::$supported = class_exists('\Propel') && version_compare(\Propel::VERSION, '1.7.0', '>=') >= 0;
        }

        return static::$supported;
    }

    /**
     * (non-PHPdoc)
     * @see \NTLAB\Report\Data\Data::canHandle()
     */
    public function canHandle($source)
    {
        return $this->isSupported() && class_exists($source.'Query');
    }

    /**
     * (non-PHPdoc)
     * @see \NTLAB\Report\Data\Data::addCondition()
     */
    public function addCondition(Parameter $parameter)
    {
        $query = $this->getQuery();
        foreach ($this->convertParameter($parameter) as $params) {
            list($column, $operator, $value) = $params;
            $this->applyQuery($query, $column, 'filterBy%s', array($value, $operator));
        }
    }

    /**
     * (non-PHPdoc)
     * @see \NTLAB\Report\Data\Data::addOrder()
     */
    public function addOrder($column, $direction)
    {
        $query = $this->getQuery();
        $this->applyQuery($query, $column, 'orderBy%s', array($direction));
    }

    /**
     * (non-PHPdoc)
     * @see \NTLAB\Report\Data\Data::fetch()
     */
    public function fetch()
    {
        $query = $this->getQuery();
        // merge subqueries
        for ($i = count($this->items); $i > 0; $i--) {
            $subquery = $this->items[$i - 1];
            if ($subquery != $query) {
                $subquery->endUse();
            }
        }
        // set distinct
        if ($this->isDistinct()) {
            $query->setDistinct();
        }

        return $query->find();
    }

    /**
     * Get query class name.
     *
     * @return string
     */
    protected function getQueryClass()
    {
        return $this->source.'Query';
    }

    /**
     * Get the query object.
     *
     * @return \ModelCriteria
     */
    protected function getQuery()
    {
        if (null === $this->query) {
            $class = $this->getQueryClass();
            $this->query = new $class();
        }

        return $this->query;
    }

    /**
     * Apply query to criteria object.
     *
     * @param \ModelCriteria $query  The criteria object
     * @param string $column  The column name, can be nested using MyModel.MyOtherModel.MyMethod
     * @param string $method_format  The method format to be called
     * @param array $args  The method arguments
     * @throws \RuntimeException
     * @return \NTLAB\Report\Data\Propel
     */
    protected function applyQuery($query, $column, $method_format, $args)
    {
        $extra = null;
        if (false !== strpos($column, self::COLUMN_CONCATENATOR)) {
            list ($column, $extra) = explode(self::COLUMN_CONCATENATOR, $column, 2);
            $method = 'use'.$column.'Query';
        } else {
            $method = sprintf($method_format, $column);
        }
        if (!is_callable(array($query, $method))) {
            throw new \RuntimeException('Method "'.$method.'" doesn\'t exist in "'.get_class($query).'".');
        }
        if (null !== $extra) {
            $subquery = $query->$method(null, \Criteria::LEFT_JOIN);
            if (!in_array($subquery, $this->items)) {
                $this->items[] = $subquery;
            }
            $this->applyQuery($subquery, $extra, $method_format, $args);
        } else {
            call_user_func_array(array($query, $method), $args);
        }

        return $this;
    }

    /**
     * (non-PHPdoc)
     * @see \NTLAB\Report\Data\Data::getColumn()
     */
    public function getColumn($column)
    {
        $result = null;
        if (false !== strpos($column, self::COLUMN_CONCATENATOR)) {
            $cols = explode(self::COLUMN_CONCATENATOR, $column);
            $colName = array_pop($cols);
            $peer = array_pop($cols).'Peer';
        } else {
            $peer = $this->report->getSource().'Peer';
            $colName = $column;
        }
        foreach (array(\BasePeer::TYPE_PHPNAME, \BasePeer::TYPE_FIELDNAME) as $type) {
            try {
                $result = call_user_func(array($peer, 'translateFieldName'), $colName, $type, \BasePeer::TYPE_COLNAME);
                break;
            } catch (\Exception $e) {
            }
        }

        return $result;
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

        switch ($parameter->getType())
        {
            case 'static':
                if ($value === 'NULL') {
                    $operator = \Criteria::ISNULL;
                    $value = null;
                } else if ($value === 'NOT_NULL') {
                    $operator = \Criteria::ISNOTNULL;
                    $value = null;
                }
                break;

            case 'date':
                $rcolumn = $parameter->getRealColumn();
                $dateType = $parameter->getDateTypeValue();
                $adapter = \Propel::getDB();
                if (Date::DATE === $dateType) {
                    $value = date($adapter->getDateFormatter(), $value);
                } else if (Date::MONTH === $dateType) {
                    $value = date('Y-m', $value);
                } else if (Date::YEAR === $dateType) {
                    $value = date('Y', $value);
                }
                $value = sprintf('%1$s %2$s %4$s%3$s%4$s',
                    $adapter->subString($rcolumn, 1, strlen($value)),
                    $operator ? $operator : \Criteria::EQUAL,
                    $value,
                    $adapter->getStringDelimiter()
                );
                $operator = \Criteria::CUSTOM;
                break;
        }
        $result[] = $this->createParam($column, $operator, $value);
    }
}