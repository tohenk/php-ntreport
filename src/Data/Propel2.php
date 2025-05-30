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

namespace NTLAB\Report\Data;

use NTLAB\Report\Report;
use NTLAB\Report\Parameter\Parameter;
use NTLAB\Report\Parameter\Date as DateParameter;
use NTLAB\Report\Parameter\DateOnly as DateOnlyParameter;
use NTLAB\Report\Parameter\DateRange as DateRangeParameter;
use NTLAB\Report\Parameter\DateMonth as DateMonthParameter;
use NTLAB\Report\Parameter\DateYear as DateYearParameter;
use NTLAB\Report\Parameter\Statix as StaticParameter;
use Propel\Runtime\ActiveQuery\Criteria;
use Propel\Runtime\Map\TableMap;

class Propel2 extends Data
{
    public const COLUMN_CONCATENATOR = '.';

    /**
     * @var \Propel\Runtime\ActiveQuery\ModelCriteria
     */
    protected $query = null;

    /**
     * @var \Propel\Runtime\ActiveQuery\ModelCriteria[]
     */
    protected $items = [];

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
        if (null === static::$supported) {
            static::$supported = class_exists('\Propel\Runtime\Propel') && version_compare(\Propel\Runtime\Propel::VERSION, '2.0.0', '>=') >= 0;
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
        $this->items = [];
        $query = $this->getQuery();
        foreach ($this->convertParameter($parameter) as $params) {
            list($column, $operator, $value) = $params;
            $this->applyQuery($query, $column, 'filterBy%s', [$value, $operator]);
        }
        $this->endQueryUses();
    }

    /**
     * (non-PHPdoc)
     * @see \NTLAB\Report\Data\Data::addGroupBy()
     */
    public function addGroupBy($column)
    {
        $this->items = [];
        $query = $this->getQuery();
        $this->applyQuery($query, $column, 'groupBy%s', []);
        $this->endQueryUses();
    }

    /**
     * (non-PHPdoc)
     * @see \NTLAB\Report\Data\Data::addOrder()
     */
    public function addOrder($column, $direction, $format = null)
    {
        $this->items = [];
        $query = $this->getQuery();
        if ($format) {
            switch ($direction) {
                case Report::ORDER_ASC:
                    $this->applyQuery($query, $column, 'addAscendingOrderByColumn', [], $format, true);
                    break;
                case Report::ORDER_DESC:
                    $this->applyQuery($query, $column, 'addDescendingOrderByColumn', [], $format, true);
                    break;
            }
        } else {
            $this->applyQuery($query, $column, 'orderBy%s', [$direction]);
        }
        $this->endQueryUses();
    }

    /**
     * (non-PHPdoc)
     * @see \NTLAB\Report\Data\Data::count()
     */
    public function count()
    {
        $query = $this->getQuery();
        // set distinct
        if ($this->isDistinct()) {
            $query->setDistinct();
        }

        return $query->count();
    }

    /**
     * (non-PHPdoc)
     * @see \NTLAB\Report\Data\Data::fetch()
     */
    public function fetch()
    {
        $query = $this->getQuery();
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
     * @return \Propel\Runtime\ActiveQuery\ModelCriteria
     */
    protected function getQuery()
    {
        if (null === $this->query) {
            $class = $this->getQueryClass();
            $this->query = $class::create();
        }

        return $this->query;
    }

    /**
     * Apply query to criteria object.
     *
     * @param \Propel\Runtime\ActiveQuery\ModelCriteria $query  The criteria object
     * @param string $column  The column name, can be nested using MyModel.MyOtherModel.MyMethod
     * @param string $method_format  The method format to be called
     * @param array $args  The method arguments
     * @param string $column_format  The format to be applied to column
     * @param bool $translate  Perform column name translation
     * @throws \RuntimeException
     * @return \NTLAB\Report\Data\Propel2
     */
    protected function applyQuery($query, $column, $method_format, $args, $column_format = null, $translate = null)
    {
        $extra = null;
        if (false !== strpos($column, self::COLUMN_CONCATENATOR)) {
            list($column, $extra) = explode(self::COLUMN_CONCATENATOR, $column, 2);
            $method = 'use'.$column.'Query';
        } elseif (false !== strpos($method_format, '%s')) {
            $method = sprintf($method_format, $column);
        } else {
            $method = $method_format;
            if ($translate) {
                $tableMap = $this->getQuery()->getTableMap();
                foreach ([TableMap::TYPE_PHPNAME, TableMap::TYPE_FIELDNAME] as $type) {
                    try {
                        $translatedColumn = $tableMap->translateFieldName($column, $type, TableMap::TYPE_COLNAME);
                        $column = $translatedColumn;
                        break;
                    } catch (\Exception $e) {
                        error_log($e->getMessage());
                    }
                }
            }
            array_unshift($args, $this->applyFormat($column_format, $column));
        }
        if (!is_callable([$query, $method])) {
            throw new \RuntimeException('Method "'.$method.'" doesn\'t exist in "'.get_class($query).'".');
        }
        if (null !== $extra) {
            $subquery = $query->$method(null, Criteria::LEFT_JOIN);
            if (!in_array($subquery, $this->items)) {
                $this->items[] = $subquery;
            }
            $this->applyQuery($subquery, $extra, $method_format, $args, $column_format, $translate);
        } else {
            call_user_func_array([$query, $method], $args);
        }

        return $this;
    }

    /**
     * End all sub query uses and merge with main query.
     *
     * @return \Propel\Runtime\ActiveQuery\ModelCriteria
     */
    protected function endQueryUses()
    {
        $query = $this->getQuery();
        // merge subqueries
        for ($i = count($this->items); $i > 0; $i--) {
            $subquery = $this->items[$i - 1];
            if ($subquery != $query) {
                $subquery->endUse();
            }
        }

        return $query;
    }

    /**
     * (non-PHPdoc)
     * @see \NTLAB\Report\Data\Data::getColumn()
     */
    public function getColumn($column)
    {
        $result = null;
        $tableMap = $this->getQuery()->getTableMap();
        if (false !== strpos($column, self::COLUMN_CONCATENATOR)) {
            $cols = explode(self::COLUMN_CONCATENATOR, $column);
            $colName = array_pop($cols);
        } else {
            $colName = $column;
        }
        foreach ([TableMap::TYPE_PHPNAME, TableMap::TYPE_FIELDNAME] as $type) {
            try {
                $result = $tableMap->translateFieldName($colName, $type, TableMap::TYPE_COLNAME);
                break;
            } catch (\Exception $e) {
            }
        }

        return $result;
    }

    /**
     * Format a date as Propel Criteria.
     *
     * @param string $column  Column name
     * @param string $dateType  Date type
     * @param int $dateValue  Date timestamp value
     * @param string $operator  Criteria operator
     * @return string
     */
    protected function formatDate($column, $dateType, $dateValue, $operator = null)
    {
        $tableMap = $this->getQuery()->getTableMap();
        $adapter = \Propel\Runtime\Propel::getAdapter($tableMap::DATABASE_NAME);
        if (DateParameter::DATE === $dateType) {
            $dateValue = date($adapter->getDateFormatter(), $dateValue);
        } elseif (DateParameter::MONTH === $dateType) {
            $dateValue = date('Y-m', $dateValue);
        } elseif (DateParameter::YEAR === $dateType) {
            $dateValue = date('Y', $dateValue);
        }

        return sprintf(
            '%1$s %2$s %4$s%3$s%4$s',
            $adapter->subString($column, 1, strlen($dateValue)),
            $operator ? $operator : Criteria::EQUAL,
            $dateValue,
            $adapter->getStringDelimiter()
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
        switch ($parameter->getType()) {
            case StaticParameter::ID:
                if ($value === 'NULL') {
                    $operator = Criteria::ISNULL;
                    $value = null;
                } elseif ($value === 'NOT_NULL') {
                    $operator = Criteria::ISNOTNULL;
                    $value = null;
                }
                break;
            case DateParameter::ID:
            case DateOnlyParameter::ID:
            case DateMonthParameter::ID:
            case DateYearParameter::ID:
                /** @var \NTLAB\Report\Parameter\Date $parameter */
                $value = $this->formatDate($parameter->getRealColumn(), $parameter->getDateTypeValue(), $value, $operator);
                $operator = Criteria::CUSTOM;
                break;
            case DateRangeParameter::ID:
                /** @var \NTLAB\Report\Parameter\DateRange $parameter */
                $value = sprintf(
                    '(%s AND %s)',
                    $this->formatDate($parameter->getRealColumn(), $parameter->getDateTypeValue(), $value, '>='),
                    $this->formatDate($parameter->getRealColumn(), $parameter->getDateTypeValue(), $parameter->getCurrentValue2(), '<=')
                );
                $operator = Criteria::CUSTOM;
                break;
        }
        $result[] = $this->createParam($column, $operator, $value);
    }
}
