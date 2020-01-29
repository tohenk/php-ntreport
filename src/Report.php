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

namespace NTLAB\Report;

use NTLAB\Report\Config\Config;
use NTLAB\Report\Data\Data;
use NTLAB\Report\Data\Pdo as PdoData;
use NTLAB\Report\Data\Propel2 as Propel2Data;
use NTLAB\Report\Engine\Csv as CsvEngine;
use NTLAB\Report\Engine\Excel as ExcelEngine;
use NTLAB\Report\Engine\Richtext as RichtextEngine;
use NTLAB\Report\Engine\Word as WordEngine;
use NTLAB\Report\Form\FormBuilder;
use NTLAB\Report\Listener\ListenerInterface;
use NTLAB\Report\Parameter\Parameter;
use NTLAB\Report\Parameter\Boolean as BooleanParameter;
use NTLAB\Report\Parameter\Checklist as ChecklistParameter;
use NTLAB\Report\Parameter\Date as DateParameter;
use NTLAB\Report\Parameter\DateOnly as DateOnlyParameter;
use NTLAB\Report\Parameter\DateRange as DateRangeParameter;
use NTLAB\Report\Parameter\DateMonth as DateMonthParameter;
use NTLAB\Report\Parameter\DateYear as DateYearParameter;
use NTLAB\Report\Parameter\Reference as ReferenceParameter;
use NTLAB\Report\Parameter\Statix as StaticParameter;
use NTLAB\Report\Validator\Validator;
use NTLAB\Report\Script\ProviderReport;
use NTLAB\Report\Script\ReportCore;
use NTLAB\Script\Core\Script;
use NTLAB\Script\Core\Manager;

// register base engines
CsvEngine::create()->register();
ExcelEngine::create()->register();
RichtextEngine::create()->register();
WordEngine::create()->register();

// register report data
PdoData::create()->register();
Propel2Data::create()->register();

// register parameters
BooleanParameter::create()->register();
ChecklistParameter::create()->register();
DateParameter::create()->register();
DateOnlyParameter::create()->register();
DateRangeParameter::create()->register();
DateMonthParameter::create()->register();
DateYearParameter::create()->register();
ReferenceParameter::create()->register();
StaticParameter::create()->register();

// register report script
Manager::addProvider(ProviderReport::getInstance());

abstract class Report
{
    const ID = 'none';

    const ORDER_ASC = 'ASC';
    const ORDER_DESC = 'DESC';

    const STATUS_OK = 0;
    const STATUS_ERR_TMPL = 1;
    const STATUS_ERR_TMPL_INVALID = 2;
    const STATUS_ERR_NO_DATA = 3;
    const STATUS_ERR_INTERNAL = 4;

    /**
     * @var \DOMDocument
     */
    protected $doc = null;

    /**
     * @var \DOMXPath
     */
    protected $xpath = null;

    /**
     * @var string
     */
    protected $title = null;

    /**
     * @var string
     */
    protected $category = null;

    /**
     * @var string
     */
    protected $source = null;

    /**
     * @var boolean
     */
    protected $distinct = null;

    /**
     * @var array
     */
    protected $parameters = array();

    /**
     * @var array
     */
    protected $orders = array();

    /**
     * @var array
     */
    protected $groups = array();

    /**
     * @var array
     */
    protected $validators = array();

    /**
     * @var array
     */
    protected $configs = array();

    /**
     * @var \NTLAB\Report\Data\Data
     */
    protected $data = null;

    /**
     * @var \NTLAB\Report\Form\FormInterface
     */
    protected $form = null;

    /**
     * @var array
     */
    protected $result = null;

    /**
     * @var boolean
     */
    protected $hasTemplate = true;

    /**
     * @var string
     */
    protected $template = null;

    /**
     * @var string
     */
    protected $templateContent = null;

    /**
     * @var \NTLAB\Script\Core\Script
     */
    protected $script = null;

    /**
     * @var mixed
     */
    protected $object = null;

    /**
     * @var int
     */
    protected $status = null;

    /**
     * @var \Exception
     */
    protected $error = null;

    /**
     * @var array
     */
    protected static $engines = array();

    /**
     * @var \NTLAB\Report\Listener\ListenerInterface[]
     */
    protected static $listeners = array();

    /**
     * @var \NTLAB\Report\Form\FormBuilder
     */
    protected static $formBuilder;

    /**
     * Add report engine.
     *
     * @param string $engine  The engine id
     * @param string $class  The report class
     */
    public static function addEngine($engine, $class)
    {
        if (!isset(static::$engines[$engine])) {
            static::$engines[$engine] = $class;
        }
    }

    /**
     * Register report listener.
     *
     * @param \NTLAB\Report\Listener\ListenerInterface $listener  Report listener
     */
    public static function addListener(ListenerInterface $listener)
    {
        if (!in_array($listener, static::$listeners)) {
            static::$listeners[] = $listener;
        }
    }

    /**
     * Get report listeners.
     *
     * @return \NTLAB\Report\Listener\ListenerInterface
     */
    public static function getListeners()
    {
        return static::$listeners;
    }

    /**
     * Get engine class.
     *
     * @param string $type  Engine id
     * @return string
     */
    public static function getEngine($type)
    {
        if (isset(static::$engines[$type])) {
            return static::$engines[$type];
        }
    }

    /**
     * Set report form builder.
     *
     * @param \NTLAB\Report\Form\FormBuilder $builder  The form builder
     */
    public static function setFormBuilder(FormBuilder $builder)
    {
        static::$formBuilder = $builder;
    }

    /**
     * Get report form builder.
     *
     * @return \NTLAB\Report\Form\FormBuilder
     */
    public static function getFormBuilder()
    {
        return static::$formBuilder;
    }

    /**
     * Load a report xml and create the instance.
     *
     * @param string $xml  XML string
     * @return \NTLAB\Report\Report
     */
    public static function load($xml)
    {
        $doc = new \DOMDocument();
        if ($doc->loadXml($xml)) {
            $xpath = new \DOMXPath($doc);
            if ($nodes = $xpath->query("//NTReport/Params/Param[@name='type']")) {
                $type = $nodes->item(0)->attributes->getNamedItem('value')->nodeValue;
                if (($class = static::getEngine($type)) && class_exists($class)) {
                    $report = new $class();
                    $report->initialize($doc, $xpath);

                    return $report;
                }
            }
        }
    }

    /**
     * Create instance.
     *
     * @return \NTLAB\Report\Report
     */
    public static function create()
    {
        return new static();
    }

    /**
     * Register report engine.
     */
    public function register()
    {
        $this->addEngine(static::ID, get_class($this));
    }

    /**
     * Constructor.
     *
     * @param \DOMDocument $doc  The report document
     * @param \DOMXPath $xpath  The xpath object
     */
    public function initialize(\DOMDocument $doc, \DOMXPath $xpath)
    {
        $this->doc = $doc;
        $this->xpath = $xpath;
        // initialize report
        $this->doInitialize();
        // report base parameter
        $nodes = $this->xpath->query('//NTReport/Params/Param');
        foreach ($nodes as $node) {
            switch ($this->nodeAttr($node, 'name')) {
                case 'type':
                    break;

                case 'title':
                    $this->title = $this->nodeAttr($node, 'value');
                    break;

                case 'category':
                    $this->category = $this->nodeAttr($node, 'value');
                    break;

                case 'model':
                case 'source':
                    $this->source = $this->nodeAttr($node, 'value');
                    $this->buildSourceParams($node->childNodes);
                    break;

                case 'distinct':
                    $this->distinct = (bool) $this->nodeAttr($node, 'value');
                    break;

                case 'validator':
                    $this->buildValidators($node->childNodes);
                    break;

                case 'config':
                    $this->buildConfigs($node->childNodes);
                    break;
            }
        }
        // specific report configuration
        $nodes = $this->xpath->query('//NTReport/Configuration');
        $this->configure($nodes->item(0)->childNodes);
    }

    /**
     * Report initialization.
     */
    protected function doInitialize()
    {
    }

    /**
     * Get script object.
     *
     * @return \NTLAB\Script\Core\Script
     */
    public function getScript()
    {
        if (null == $this->script) {
            $this->script = new Script();
        }

        return $this->script;
    }

    /**
     * Get report object.
     *
     * @return mixed the report object
     */
    public function getObject()
    {
        return $this->object;
    }

    /**
     * Set the report object.
     *
     * @param mixed $v  The report object
     * @return \NTLAB\Report\Report
     */
    public function setObject($v)
    {
        $this->object = $v;

        return $this;
    }

    /**
     * Get DOMNode attribute value.
     *
     * @param \DOMNode $node  The node
     * @param string $attr  The attribute
     * @param mixed $default  Node default value
     * @return string
     */
    protected function nodeAttr(\DOMNode $node, $attr, $default = null)
    {
        if ($node->hasAttributes() && ($nodettr = $node->attributes->getNamedItem($attr))) {
            return $nodettr->nodeValue;
        }

        return $default;
    }

    /**
     * Configure report.
     *
     * @param \DOMNodeList $nodes  The configuration nodes
     */
    abstract protected function configure(\DOMNodeList $nodes);

    /**
     * Add model parameter.
     *
     * @param \DOMNode $node  The node parameter
     */
    protected function addSourceParameter(\DOMNode $node)
    {
        $type = $this->nodeAttr($node, 'type', 'static');
        if (null === ($class = Parameter::getHandler($type))) {
            return;
        }
        $name = $this->nodeAttr($node, 'name');
        $options = array(
            'name' => $name,
            'column' => $this->nodeAttr($node, 'column'),
            'value' => $this->nodeAttr($node, 'value'),
            'is_default' => (bool) $this->nodeAttr($node, 'default', false),
            'is_changeable' => (bool) $this->nodeAttr($node, 'change', true)
        );
        if ($operator = $this->nodeAttr($node, 'operator')) {
            $options['operator'] = sprintf(' %s ', trim($operator));
        }
        if ($label = $this->nodeAttr($node, 'label')) {
            $options['label'] = $label;
        }
        if ($onValue = $this->nodeAttr($node, 'getvalue')) {
            $options['on_value'] = $onValue;
        }
        if (($parent = $this->nodeAttr($node, 'parent')) && isset($this->parameters[$parent])) {
            $options['parent'] = $this->parameters[$parent];
        }
        $parameter = new $class();
        $parameter->initialize($type, $this, $options);
        $this->parameters[$name] = $parameter;
    }

    /**
     * Add model grouping.
     *
     * @param \DOMNode $node  The node order
     */
    protected function addSourceGrouping(\DOMNode $node)
    {
        $column = $this->nodeAttr($node, 'name');
        $this->groups[] = $column;
    }

    /**
     * Add model sorting.
     *
     * @param \DOMNode $node  The node order
     */
    protected function addSourceOrdering(\DOMNode $node)
    {
        $column = $this->nodeAttr($node, 'name');
        $dir = $this->nodeAttr($node, 'dir', static::ORDER_ASC);
        $format = $this->nodeAttr($node, 'fmt');
        $this->orders[] = array($column, $dir, $format);
    }

    /**
     * Build model parameters.
     *
     * @param \DOMNodeList $nodes  The node list
     */
    protected function buildSourceParams(\DOMNodeList $nodes)
    {
        foreach ($nodes as $node) {
            switch (strtolower($node->nodeName)) {
                case 'param':
                    $this->addSourceParameter($node);
                    break;

                case 'group':
                    $this->addSourceGrouping($node);
                    break;

                case 'order':
                    $this->addSourceOrdering($node);
                    break;
            }
        }
    }

    /**
     * Build report validators.
     *
     * @param \DOMNodeList $nodes  The node list
     */
    protected function buildValidators(\DOMNodeList $nodes)
    {
        foreach ($nodes as $node) {
            if (($name = $this->nodeAttr($node, 'name')) && ($expr = $this->nodeAttr($node, 'expr'))) {
                $validator = new Validator($name, $expr);
                $validator->setReport($this);
                $this->validators[$name] = $validator;
            }
        }
    }

    /**
     * Build report configs.
     *
     * @param \DOMNodeList $nodes  The node list
     */
    protected function buildConfigs(\DOMNodeList $nodes)
    {
        foreach ($nodes as $node) {
            if (($name = $this->nodeAttr($node, 'name')) && ($model = $this->nodeAttr($node, 'model')) && ($column = $this->nodeAttr($node, 'column'))) {
                $attributes = array();
                foreach ($node->attributes as $attr) {
                    if (!in_array($attr->nodeName, array('name', 'model', 'column', 'label', 'help', 'default', 'object', 'required'))) {
                        $attributes[$attr->nodeName] = $attr->nodeValue;
                    }
                }
                $config = new Config($name, $model, $column, $this->nodeAttr($node, 'label'), $this->nodeAttr($node, 'help'), $this->nodeAttr($node, 'object'), $this->nodeAttr($node, 'required'), $attributes);
                $config->setReport($this);
                $this->configs[$name] = $config;
            }
        }
    }

    /**
     * Build report form parameter.
     *
     * @param array $defaults Default values
     */
    protected function buildForm($defaults)
    {
        ReportCore::setReport($this);
        static::$formBuilder->setReport($this);
        $this->form = static::$formBuilder->build($defaults);
    }

    /**
     * Get report title.
     *
     * @return string
     */
    public function getTitle()
    {
        return $this->title;
    }

    /**
     * Get report category.
     *
     * @return string
     */
    public function getCategory()
    {
        return $this->category;
    }

    /**
     * Get report model class.
     *
     * @return string
     */
    public function getSource()
    {
        return $this->source;
    }

    /**
     * Get the report data source.
     *
     * @return \NTLAB\Report\Data\Data
     */
    public function getReportData()
    {
        if (null === $this->data) {
            if ($data = Data::getHandler($this->getSource())) {
                $data->setReport($this);
            }
            $this->data = $data;
        }

        return $this->data;
    }

    /**
     * Get report parametes form.
     *
     *
     * @param array $defaults  Default values
     * @return \NTLAB\Report\Form\FormInterface
     */
    public function getForm($defaults = array())
    {
        if (null == $this->form) {
            $this->buildForm($defaults);
        }

        return $this->form;
    }

    /**
     * Get report parameters.
     *
     * @return \NTLAB\Report\Parameter\Parameter[]
     */
    public function getParameters()
    {
        return $this->parameters;
    }

    /**
     * Get orders column.
     *
     * @return array
     */
    public function getOrders()
    {
        return $this->orders;
    }

    /**
     * Get groups column.
     *
     * @return array
     */
    public function getGroups()
    {
        return $this->groups;
    }

    /**
     * Get report configurations.
     *
     * @return \NTLAB\Report\Config\Config[]
     */
    public function getConfigs()
    {
        return $this->configs;
    }

    /**
     * Get report template if available.
     *
     * @return string
     */
    public function getTemplate()
    {
        return $this->template;
    }

    /**
     * Save report configuration.
     *
     * @return \NTLAB\Report\Report
     */
    public function saveConfigs()
    {
        $objects = array();
        foreach ($this->configs as $name => $config) {
            $var = $name;
            $context = null;
            if ($config->getObjectExpr()) {
                $context = $this->getScript()->evaluate($config->getObjectExpr());
            } else {
                $context = $this->getObject();
            }
            $this->getScript()->getVarContext($context, $var);
            if ($context && ($handler = Manager::getContextHandler($context))) {
                $value = $config->getFormValue();
                $method = $handler->setMethod($context, $var);
                if (is_callable($method)) {
                    call_user_func($method, $value);
                    if (!in_array($context, $objects)) {
                        $objects[] = $context;
                    }
                }
            }
        }
        foreach ($objects as $object) {
            if ($handler = Manager::getContextHandler($object)) {
                $handler->flush($object);
            }
        }

        return $this;
    }

    /**
     * Get report generation status.
     *
     * @return int
     */
    public function getStatus()
    {
        return $this->status;
    }

    /**
     * Get report generation error object.
     *
     * @return \Exception
     */
    public function getError()
    {
      return $this->error;
    }

    /**
     * Get model result.
     *
     * @return array The result
     */
    protected function getResult()
    {
        if ($this->form->isValid()) {
            if (!($data = $this->getReportData())) {
                throw new \RuntimeException('No report data can handle '.$this->source);
            }
            // apply query filter
            foreach ($this->getParameters() as $param) {
                if (!$param->isSelected()) {
                    continue;
                }
                $data->addCondition($param);
            }
            // apply grouping
            foreach ($this->getGroups() as $group) {
                $data->addGroupBy($group);
            }
            // apply sorting
            foreach ($this->getOrders() as $order) {
                $data->addOrder($order[0], $order[1], $order[2]);
            }
            $data->setDistinct($this->distinct);

            return $data->fetch();
        }
    }

    /**
     * Is report has template.
     *
     * @return bool
     */
    public function hasTemplate()
    {
        return $this->hasTemplate;
    }

    /**
     * Set template content.
     *
     * @param string $content  The content
     * @return \NTLAB\Report\Report
     */
    public function setTemplateContent($content)
    {
        $this->templateContent = $content;

        return $this;
    }

    /**
     * Get template content.
     *
     * @return string
     */
    public function getTemplateContent()
    {
        return $this->templateContent;
    }

    /**
     * Check if the result is available.
     *
     * @return bool
     */
    public function hasResult()
    {
        return count($this->result) ? true : false;
    }

    abstract protected function build();

    /**
     * Generate the report.
     *
     * @see generateFromObjects()
     * @return string|bool The generated content or false if a failure happened
     */
    public function generate()
    {
        return $this->generateFromObjects($this->getResult());
    }

    /**
     * Generate the report from supplied objects.
     *
     * @param array $objects  The objects
     * @return string|bool The generated content or false if a failure happened
     */
    public function generateFromObjects($objects)
    {
        ReportCore::setReport($this);

        $this->error = null;
        $this->status = null;
        $this->result = $objects;
        if (count($this->result)) {
            if ($this->hasTemplate() && !$this->templateContent) {
                $this->status = static::STATUS_ERR_TMPL;
            } else {
                try {
                    if (null !== ($content = $this->build())) {
                        $this->status = static::STATUS_OK;
                        return $content;
                    }
                } catch (\Exception $e) {
                    $this->error = $e;
                    error_log($this->getExceptionMessage($e));
                }
                $this->status = static::STATUS_ERR_INTERNAL;
            }
        } else {
            $this->status = static::STATUS_ERR_NO_DATA;
        }

        return false;
    }

    protected function getExceptionMessage(\Exception $exception, $wrapper = '%s: [%s]')
    {
        $message = null;
        while (null !== $exception) {
            if ($msg = $exception->getMessage()) {
                if (null == $message) {
                    $message = $msg;
                } else {
                    $message = sprintf($wrapper, $message, $msg);
                }
            }
            $exception = $exception->getPrevious();
        }
  
        return $message;
    }

    /**
     * Check if current report is valid by executing the validators.
     *
     * @return bool
     */
    public function isValid()
    {
        foreach ($this->validators as $validator) {
            if (!$validator->isValid()) {
                return false;
            }
        }

        return true;
    }

    /**
     * Populate report config values.
     *
     * @return \NTLAB\Report\Report
     */
    public function populateConfigValues()
    {
        foreach ($this->configs as $config) {
            $config->updateConfigValue();
        }

        return $this;
    }
}
