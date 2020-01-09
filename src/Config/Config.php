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

namespace NTLAB\Report\Config;

use NTLAB\Report\Report;

class Config
{
    protected $name = null;
    protected $model = null;
    protected $column = null;
    protected $label = null;
    protected $help = null;
    protected $required = null;
    protected $attributes = array();
    protected $report = null;
    protected $objectexpr = null;
    protected $object = null;

    /**
     * Constructor.
     *
     * @param string $name  Config name
     * @param string $model  Config model
     * @param string $column  Config column
     * @param string $label  Config label
     * @param string $help  Config help
     * @param string $objectexpr  Object expression
     * @param array $attributes  Attributes
     */
    public function __construct($name, $model, $column, $label = null, $help = null, $objectexpr = null, $required = null, $attributes = array())
    {
        $this->name = $name;
        $this->model = $model;
        $this->column = $column;
        $this->label = $label;
        $this->help = $help;
        $this->objectexpr = $objectexpr;
        $this->required = $required;
        $this->attributes = $attributes;
    }

    /**
     * Set the report object.
     *
     * @param \NTLAB\Report\Report $report  The report object
     * @return \NTLAB\Report\Config\Config
     */
    public function setReport(Report $report)
    {
        $this->report = $report;

        return $this;
    }

    /**
     * Get the name.
     *
     * @return string
     */
    public function getName()
    {
        return $this->name;
    }

    /**
     * Get the model class name.
     *
     * @return string
     */
    public function getModel()
    {
        return $this->model;
    }

    /**
     * Get the model column name.
     *
     * @return string
     */
    public function getColumn()
    {
        return $this->column;
    }

    /**
     * Get widget label.
     *
     * @return string
     */
    public function getLabel()
    {
        return $this->label;
    }

    /**
     * Get widget help text.
     *
     * @return string
     */
    public function getHelp()
    {
        return $this->help;
    }

    /**
     * Get the object script expression.
     *
     * @return string
     */
    public function getObjectExpr()
    {
        return $this->objectexpr;
    }

    /**
     * Get the config requirement state.
     *
     * @return bool
     */
    public function isRequired()
    {
        return $this->required;
    }

    /**
     * Get the node attributes.
     *
     * @return array
     */
    public function getAttributes()
    {
        return $this->attributes;
    }

    /**
     * Get the form value.
     *
     * @return mixed
     */
    public function getFormValue()
    {
        $values = $this->report->getForm()->getValue('config');

        return isset($values[$this->column]) ? $values[$this->column] : null;
    }

    /**
     * Get config default value.
     *
     * @return mixed
     */
    public function getValue()
    {
        if (null == $this->object) {
            if ($this->objectexpr) {
                $this->object = $this->report->getScript()->evaluate($this->objectexpr);
            } else {
                $this->object = $this->report->getObject();
            }
        }
        $value = null;
        if ($this->report->getScript()->getVar($value, $this->name, $this->object)) {
            return $value;
        }
    }

    /**
     * Update value.
     *
     * @return \NTLAB\Report\Config\Config
     */
    public function updateConfigValue()
    {
        if (null !== ($value = $this->getValue())) {
            $form = $this->report->getForm();
            $values = $form->getDefault('config');
            $values[$this->getColumn()] = $value;
            $form->updateValue('config', $values);
        }

        return $this;
    }
}