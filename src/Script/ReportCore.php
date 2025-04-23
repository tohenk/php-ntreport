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

namespace NTLAB\Report\Script;

use NTLAB\Script\Core\Manager;
use NTLAB\Script\Core\Module;
use NTLAB\Script\Core\Script;
use NTLAB\Report\Report;
use NTLAB\Report\Symbol;

/**
 * Report core functions.
 *
 * @author Toha <tohenk@yahoo.com>
 * @id report.core
 */
class ReportCore extends Module
{
    /**
     * @var \NTLAB\Report\Report
     */
    protected static $report = null;

    /**
     * @var \NTLAB\Report\Symbol[]
     */
    protected static $symbols = [];

    /**
     * Set report instance.
     *
     * @param \NTLAB\Report\Report $report  The report
     */
    public static function setReport(Report $report)
    {
        self::$report = $report;
    }

    /**
     * Get report object.
     *
     * @return \NTLAB\Report\Report
     */
    public static function getReport()
    {
        return static::$report;
    }

    /**
     * Get report parameter.
     *
     * @param string $name  The parameter name
     * @return \NTLAB\Report\Parameter\Parameter
     */
    public static function getParameter($name)
    {
        $parameters = self::$report->getParameters();

        return isset($parameters[$name]) ? $parameters[$name] : null;
    }

    /**
     * Get the current report parent id (usually object primary key).
     *
     * @param string $idx  The index
     * @return int
     * @func parentid
     */
    public function f_ParentId($idx = null)
    {
        $keys = null;
        foreach (Manager::getContexes() as $context) {
            if (!$context->canHandle($this->getReport()->getObject())) {
                continue;
            }
            if (null !== ($pair = $context->getKeyValuePair($this->getReport()->getObject()))) {
                $keys = $pair[0];
            }
        }
        if (is_array($keys) && null !== $idx) {
            $keys = $keys[(int) $idx];
        }

        return $keys;
    }

    /**
     * Get the report parent value.
     *
     * @param string $column  Column name
     * @return mixed
     * @func pvar
     */
    public function f_ParentVar($column)
    {
        $value = null;
        $this->getReport()->getScript()->getVar($value, $column, $this->getReport()->getObject());

        return $value;
    }

    /**
     * Check if report parameter `name` is selected.
     *
     * @param string $name  The parameter name
     * @return int
     * @func repselect
     */
    public function f_RepSelected($name)
    {
        $result = false;
        if ($param = $this->getParameter($name)) {
            $result = $param->isSelected();
        }

        return Script::asBool($result);
    }

    /**
     * Get the report parameter title named `name`.
     *
     * @param string $name  The parameter name
     * @return string
     * @func reptitle
     */
    public function f_RepTitle($name)
    {
        if ($param = $this->getParameter($name)) {
            return $param->getTitle();
        }
    }

    /**
     * Get the report parameter condition value named `name`.
     *
     * @param string $name  The parameter name
     * @return string
     * @func repcond
     */
    public function f_RepCond($name)
    {
        if ($param = $this->getParameter($name)) {
            return $param->getValue();
        }
    }

    /**
     * Get the report configuration value named `name`.
     *
     * @param string $name  The configuration name
     * @return string
     * @func repcfg
     */
    public function f_RepConfig($name)
    {
        $value = null;
        $configs = $this->getReport()->getConfigs();
        if (isset($configs[$name])) {
            $value = $configs[$name]->getFormValue();
        }

        return $value;
    }

    /**
     * Get report symbol.
     *
     * @param int $index  Symbol index
     * @return string
     * @func sym
     */
    public function f_Symbol($index)
    {
        if ($symbol = $this->getReport()->getSymbol($index)) {
            foreach (static::$symbols as $sym) {
                if ($symbol === (string) $sym) {
                    return $sym;
                }
            }
            $sym = new Symbol($symbol);
            static::$symbols[] = $sym;

            return $sym;
        }
    }
}
