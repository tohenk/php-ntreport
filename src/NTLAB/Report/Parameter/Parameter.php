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

use NTLAB\Report\Report;
use NTLAB\Script\Core\Manager;
use NTLAB\Script\Context\ArrayVar;
use Propel\Runtime\DataFetcher\DataFetcherInterface;

abstract class Parameter
{
    const ID = 'none';

    /**
     * @var \NTLAB\Report\Report
     */
    protected $report = null;

    /**
     * @var string
     */
    protected $name = null;

    /**
     * @var string
     */
    protected $type = null;

    /**
     * @var string
     */
    protected $column = null;

    /**
     * @var \NTLAB\Report\Parameter\Parameter
     */
    protected $parent = null;

    /**
     * @var string
     */
    protected $value = null;

    /**
     * @var boolean
     */
    protected $default = null;

    /**
     * @var boolean
     */
    protected $changeable = null;

    /**
     * @var string
     */
    protected $operator = null;

    /**
     * @var string
     */
    protected $label = null;

    /**
     * @var string
     */
    protected $onValue = null;

    /**
     * @var string
     */
    protected $field = null;

    /**
     * @var string
     */
    protected $fieldCheck = null;

    /**
     * @var boolean
     */
    protected $supportForm = true;

    /**
     * @var array
     */
    protected $widgets = array();

    /**
     * @var array
     */
    protected $values = null;

    /**
     * @var array
     */
    protected static $handlers = array();

    /**
     * Add parameter handler.
     *
     * @param string $type  Parameter type
     * @param string $class  Handler class
     */
    public static function addHandler($type, $class)
    {
        if (!isset(static::$handlers[$type])) {
            static::$handlers[$type] = $class;
        }
    }

    /**
     * Get handler class.
     *
     * @param string $type  Parameter type
     * @return string
     */
    public static function getHandler($type)
    {
        if (isset(static::$handlers[$type])) {
            return static::$handlers[$type];
        }
    }

    /**
     * Create instance.
     *
     * @return \NTLAB\Report\Parameter\Parameter
     */
    public static function create()
    {
        return new static();
    }

    /**
     * Constructor.
     *
     * @param string $type  Parameter type
     * @param \NTLAB\Report\Report $report  The report
     * @param array $options  Parameter options
     */
    public function initialize($type, $report, $options = array())
    {
        $this->type = $type;
        $this->report = $report;
        foreach ($options as $k => $v) {
            switch (strtolower($k)) {
                case 'name':
                    $this->name = $v;
                    break;

                case 'column':
                    $this->column = $v;
                    break;

                case 'value':
                    $this->value = $v;
                    break;

                case 'parent':
                    $this->parent = $v;
                    break;

                case 'operator':
                    $this->operator = $v;
                    break;

                case 'label':
                    $this->label = $v;
                    break;

                case 'is_default':
                    $this->default = (bool) $v;
                    break;

                case 'is_changeable':
                    $this->changeable = (bool) $v;
                    break;

                case 'on_value':
                    $this->onValue = $v;
                    break;
            }
        }
        $this->field = $this->report->getFormBuilder()->normalize($this->name);
        $this->fieldCheck = $this->field.'_check';
        $this->configure();
    }

    /**
     * Configure parameter.
     */
    protected function configure()
    {
    }

    /**
     * Get parameter identifier.
     *
     * @return string
     */
    public function getId()
    {
        return static::ID;
    }

    /**
     * Register parameter handler.
     */
    public function register()
    {
        $this->addHandler($this->getId(), get_class($this));
    }

    /**
     * Get real column name.
     *
     * @return string
     */
    public function getRealColumn()
    {
        return $this->report->getReportData()->getColumn($this->column);
    }

    /**
     * Get parameter name.
     *
     * @return string
     */
    public function getName()
    {
        return $this->name;
    }

    /**
     * Get parameter type.
     *
     * @return string
     */
    public function getType()
    {
        return $this->type;
    }

    /**
     * Get model column name.
     *
     * @return string
     */
    public function getColumn()
    {
        return $this->column;
    }

    /**
     * Get default value.
     *
     * @return string
     */
    public function getDefaultValue()
    {
        if ($this->value) {
            return $this->report->getScript()->evaluate($this->value);
        }
    }

    /**
     * Get parent.
     *
     * @return \NTLAB\Report\Parameter\Parameter
     */
    public function getParent()
    {
        return $this->parent;
    }

    /**
     * Get operator.
     *
     * @return string
     */
    public function getOperator()
    {
        return $this->operator;
    }

    /**
     * Get label.
     *
     * @return string
     */
    public function getLabel()
    {
        return $this->label;
    }

    /**
     * Get default include state.
     *
     * @return bool
     */
    public function isDefault()
    {
        return $this->default;
    }

    /**
     * Get deselected state.
     *
     * @return bool
     */
    public function isChangeable()
    {
        return $this->changeable;
    }

    /**
     * Get parameter field name.
     *
     * @return string
     */
    public function getFieldName()
    {
        return $this->field;
    }

    /**
     * Get parameter field check name.
     *
     * @return string
     */
    public function getCheckName()
    {
        return $this->fieldCheck;
    }

    /**
     * Get actual parameter value.
     *
     * @return string
     */
    public function isSelected()
    {
        if (false === $this->isChangeable()) {
            return true;
        }

        return $this->getFormValue($this->fieldCheck);
    }

    /**
     * Check if parameter has support of form widget.
     *
     * @return bool
     */
    public function isSupportForm()
    {
        return $this->supportForm;
    }

    /**
     * Get form widgets associated with the parameter.
     *
     * @return array
     */
    public function getWidgets()
    {
        return $this->widgets;
    }

    /**
     * Add widget as part of the parameter.
     *
     * @param string $name  Widget name
     * @return \NTLAB\Report\Parameter\Parameter
     */
    public function addWidget($name)
    {
        if (!in_array($name, $this->widgets)) {
            $this->widgets[] = $name;
        }

        return $this;
    }

    /**
     * Get form parameter value.
     *
     * @param string $field  The form field
     * @return mixed
     */
    public function getFormValue($field)
    {
        $form = $this->report->getForm();
        $value = null;
        // value from binded data
        $values = $form->getValue('param');
        if (isset($values[$field])) {
            $value = $values[$field];
        } else {
            // try default value
            $defaults = $form->getDefault('param');
            if (isset($defaults[$field])) {
                $value = $defaults[$field];
            }
        }

        return $value;
    }

    /**
     * Get actual parameter value.
     *
     * @return string
     */
    public function getValue()
    {
        return $this->getFormValue($this->field);
    }

    /**
     * Get report title.
     *
     * @return string
     */
    public function getTitle()
    {
        $title = $this->getDefaultTitle();
        foreach ($this->report->getListeners() as $listener) {
            if ($listener->formatTitle($this, $title)) {
                break;
            }
        }

        return $title;
    }

    /**
     * Get default title.
     *
     * @return string
     */
    protected function getDefaultTitle()
    {
        return $this->getName();
    }

    /**
     * Get evaluated value.
     *
     * @return string
     */
    protected function evalValue($value)
    {
        if (strlen($value) && strlen($this->onValue)) {
            $script = $this->report->getScript();
            $var = new ArrayVar(array('value' => $value));
            $script->setContext($var);
            $value = $script->evaluate($this->onValue);
        }

        return $value;
    }

    /**
     * Get evaluated parameter value.
     *
     * @return string
     */
    public function getCurrentValue()
    {
        return $this->evalValue($this->getValue());
    }

    /**
     * Get list of values of default value.
     *
     * @return array
     */
    public function getValues()
    {
        if (null == $this->values) {
            $this->values = array();
            try {
                if (($value = $this->getDefaultValue()) &&
                    (
                        is_array($value) ||
                        ($value instanceof \ArrayObject) ||
                        ($value instanceof \ArrayAccess) ||
                        ($value instanceof ArrayVar) ||
                        ($value instanceof DataFetcherInterface)
                    )) {
                    foreach ($value as $ref) {
                        if ($handler = Manager::getContextHandler($ref)) {
                            if (null !== ($pair = $handler->getKeyValuePair($ref))) {
                                list($k, $v) = $pair;
                                $this->values[$k] = $v;
                            }
                        }
                    }
                }
            } catch (\Exception $e) {
            }
        }

        return $this->values;
    }

    public function __toString()
    {
        return sprintf('%s: %s (%s)', strtoupper($this->type), $this->name, $this->column);
    }
}