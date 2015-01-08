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

namespace NTLAB\Report\Util\Excel;

use NTLAB\Script\Core\Script;

class Filer
{
    const BAND_TITLE = 'title';
    const BAND_HEADER = 'header';
    const BAND_MASTER_DATA = 'master-data';
    const BAND_FOOTER = 'footer';
    const BAND_SUMMARY = 'summary';

    /**
     * @var string
     */
    protected $template = null;

    /**
     * @var string
     */
    protected $sheet = null;

    /**
     * @var array
     */
    protected $bands = array();

    /**
     * @var array
     */
    protected $datas = array();

    /**
     * @var array
     */
    protected $allBands = array(
        self::BAND_TITLE,
        self::BAND_HEADER,
        self::BAND_MASTER_DATA,
        self::BAND_FOOTER,
        self::BAND_SUMMARY
    );

    /**
     * @var \NTLAB\Script\Core\Script
     */
    protected $script = null;

    /**
     * @var array
     */
    protected $objects = null;

    /**
     * @var string
     */
    protected $fieldSign = '$';

    /**
     * @var boolean
     */
    protected $autoFit = true;

    /**
     * @var float
     */
    protected $rowHeight = null;

    /**
     * @var string
     */
    protected $defaultWriter = null;

    /**
     * @var array
     */
    protected $dataCells = array();

    /**
     * Get excel template.
     *
     * @return string
     */
    public function getTemplate()
    {
        return $this->template;
    }

    /**
     * Set excel template.
     *
     * @param string $value  Template filename
     * @return \NTLAB\Report\Util\Excel\Filer
     */
    public function setTemplate($value)
    {
        $this->template = $value;

        return $this;
    }

    /**
     * Get template sheet name.
     *
     * @return string
     */
    public function getSheet()
    {
        return $this->sheet;
    }

    /**
     * Set template sheet name.
     *
     * @param string $value  Sheet name
     * @return \NTLAB\Report\Util\Excel\Filer
     */
    public function setSheet($value)
    {
        $this->sheet = $value;

        return $this;
    }

    /**
     * Get field signature.
     *
     * @return string
     */
    public function getFieldSign()
    {
        return $this->fieldSign;
    }

    /**
     * Set field signature.
     *
     * @param string $signature  The signature
     * @return \NTLAB\Report\Util\Excel\Filer
     */
    public function setFieldSign($signature)
    {
        $this->fieldSign = $signature;

        return $this;
    }

    /**
     * Get auto fit data row.
     *
     * @return boolean
     */
    public function getAutoFit()
    {
        return $this->autoFit;
    }

    /**
     * Set data row autofit.
     *
     * @param boolean $value  Auto fit value
     * @return \NTLAB\Report\Util\Excel\Filer
     */
    public function setAutoFit($value)
    {
        $this->autoFit = (bool) $value;

        return $this;
    }

    /**
     * Get data row height.
     *
     * @return float
     */
    public function getRowHeight()
    {
        return $this->rowHeight;
    }

    /**
     * Set data row height.
     *
     * @param float $value  The row height
     * @return \NTLAB\Report\Util\Excel\Filer
     */
    public function setRowHeight($value)
    {
        $this->rowHeight = $value;

        return $this;
    }

    /**
     * Get output writer.
     *
     * @return string
     */
    public function getDefaultWriter()
    {
        return $this->defaultWriter;
    }

    /**
     * Set output writer.
     *
     * @param string $value  The writer class
     * @return \NTLAB\Report\Util\Excel\Filer
     */
    public function setDefaultWriter($value)
    {
        $this->defaultWriter = $value;

        return $this;
    }

    /**
     * Add report band range.
     *
     * @param string $band  The band name
     * @param string $range  The range
     * @return \NTLAB\Report\Util\Excel\Filer
     */
    public function addBand($band, $range)
    {
        if (in_array($band, $this->allBands)) {
            $this->bands[$band] = $range;
        }

        return $this;
    }

    /**
     * Add report data.
     *
     * @param string $field  The field name
     * @param string $expression  Value expression
     * @return \NTLAB\Report\Util\Excel\Filer
     */
    public function addData($field, $expression)
    {
        $this->datas[$field] = $expression;

        return $this;
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
     * Set the script object.
     *
     * @param \NTLAB\Script\Core\Script $script  The script object
     * @return \NTLAB\Report\Util\Excel\Filer
     */
    public function setScript(Script $script)
    {
        $this->script = $script;

        return $this;
    }

    /**
     * Copy worksheet property.
     *
     * @param \PHPExcel_Worksheet $source  Source worksheet
     * @param \PHPExcel_Worksheet $dest  Destination worksheet
     * @return \NTLAB\Report\Util\Excel\Filer
     */
    protected function copySheetProp(\PHPExcel_Worksheet $source, \PHPExcel_Worksheet $dest)
    {
        $dest->setTitle($source->getTitle());
        $dest->setPageSetup(clone $source->getPageSetup());
        $dest->setPageMargins(clone $source->getPageMargins());
        $dest->setSheetView(clone $source->getSheetView());

        return $this;
    }

    /**
     * Get merged cell.
     *
     * @param \PHPExcel_Worksheet $sheet  The worksheet object
     * @param string $cell  The cell
     * @return string
     */
    protected function getCellMerged(\PHPExcel_Worksheet $sheet, $cell)
    {
        foreach ($sheet->getMergeCells() as $key => $range) {
            list($rangeStart, $rangeEnd) = explode(':', $key);
            if ($rangeStart == $cell) {
                return $range;
            }
        }
    }

    /**
     * Merge cells.
     *
     * @param \PHPExcel_Worksheet $sheet  The worksheet
     * @param string $cell  The origin cell
     * @param int $width  The number of columns from origin to merge
     * @param int $height  The number of rows from origin to merge
     * @return \NTLAB\Report\Util\Excel\Filer
     */
    protected function mergeCells(\PHPExcel_Worksheet $sheet, $cell, $width, $height)
    {
        list($rangeStart, $rangeEnd) = \PHPExcel_Cell::rangeBoundaries($cell);
        $rangeEnd[0] = (int) $rangeStart[0] + $width - 1;
        $rangeEnd[1] = (int) $rangeStart[1] + $height - 1;
        $range = \PHPExcel_Cell::stringFromColumnIndex($rangeStart[0] - 1).$rangeStart[1].':'.\PHPExcel_Cell::stringFromColumnIndex($rangeEnd[0] - 1).$rangeEnd[1];
        $sheet->mergeCells($range);

        return $this;
    }

    /**
     * Replace the tag.
     *
     * @param string $tag  The value to replace the tag of
     * @return string
     */
    protected function replaceTag($tag)
    {
        preg_match_all('/%([^%]+)%/', $tag, $matches, PREG_PATTERN_ORDER);
        foreach ($matches[1] as $match) {
            $replacement = $this->getScript()->evaluate($match);
            $tag = str_replace('%'.$match.'%', $replacement, $tag);
        }

        return $tag;
    }

    /**
     * Copy a cell range from a sheet to another.
     *
     * @param \PHPExcel_Worksheet $source  The source worksheet
     * @param \PHPExcel_Worksheet $dest  The destination worksheet
     * @param string $range  The source range
     * @param array $anchor  The destination anchor
     * @param bool $replaceTag  True to replace tag before applying value to destination sheet
     * @return array Range boundaries array(start, end)
     */
    protected function copyRange(\PHPExcel_Worksheet $source, \PHPExcel_Worksheet $dest, $range, &$anchor, $replaceTag = true)
    {
        list($rangeStart, $rangeEnd) = \PHPExcel_Cell::rangeBoundaries($range);
        if (null == $anchor) {
            $anchor = $rangeStart;
        }
        $cols = $rangeEnd[0] - $rangeStart[0];
        $rows = $rangeEnd[1] - $rangeStart[1];
        for ($row = 0; $row <= $rows; $row++) {
            for ($col = 0; $col <= $cols; $col++) {
                $scell = $source->getCellByColumnAndRow($rangeStart[0] + $col - 1, $rangeStart[1] + $row);
                $dcell = $dest->getCellByColumnAndRow($anchor[0] + $col - 1, $anchor[1] + $row);
                // copy value, ignore empty cell
                if ($svalue = $scell->getValue()) {
                    $dcell->setValue($replaceTag ? $this->replaceTag($svalue) : $svalue);
                }
                // copy style
                $dest->getStyle($dcell->getCoordinate())
                    ->applyFromArray(Style::styleToArray($source->getStyle($scell->getCoordinate())));
                // merge cells
                if ($merged = $this->getCellMerged($source, $scell->getCoordinate())) {
                    $dim = \PHPExcel_Cell::rangeDimension($merged);
                    $this->mergeCells($dest, $dcell->getCoordinate(), $dim[0], $dim[1]);
                }
                // set column width
                $width = $source->getColumnDimension($scell->getColumn())
                    ->getWidth();
                if (($cdim = $dest->getColumnDimension($dcell->getColumn())) && $cdim->getWidth() !== $width) {
                    $cdim->setWidth($width);
                }
                // set row height
                if ($col == 0) {
                    $height = $source->getRowDimension($scell->getRow())
                        ->getRowHeight();
                    if (($rdim = $dest->getRowDimension($dcell->getRow())) && $rdim->getRowHeight() !== $height) {
                        $rdim->setRowHeight($height);
                    }
                }
            }
        }
        $ranges = array($anchor, array($anchor[0] + $cols, $anchor[1] + $rows));
        // increment anchor rows
        $anchor[1] = $anchor[1] + $rows + 1;

        return $ranges;
    }

    /**
     * Fill range for master-data.
     *
     * @param \PHPExcel_Cell $cell  The cell
     * @param string $name  Data name
     * @param string $value  Cell value
     * @return mixed
     */
    protected function fillData(\PHPExcel_Cell $cell, $name, $value)
    {
        // is variable?
        if (!$this->getScript()->getVar($value, $name, $this->getScript()->getContext())) {
            // no, it is within data tags?
            if (array_key_exists($name, $this->datas)) {
                // yes, parse the expression
                $value = $this->datas[$name];
                if ($value) {
                    $value = $this->getScript()->evaluate($value);
                }
            }
        }
        // apply rich text to value, it's usable only for Excel2007 writer
        if ($value) {
            $value = RichText::create($value);
        }
        // assign data cell dimension for summary
        if (!array_key_exists($name, $this->dataCells)) {
            $this->dataCells[$name] = array(
                'column' => $cell->getColumn(),
                'rowStart' => $cell->getRow(),
                'rowEnd' => $cell->getRow()
            );
        } else {
            $this->dataCells[$name]['rowEnd'] = $cell->getRow();
        }

        return $value;
    }

    /**
     * Fill range for summary.
     *
     * @param \PHPExcel_Cell $cell  The cell
     * @param string $name  Data name
     * @param string $value  Cell value
     * @return mixed
     */
    protected function fillSummary(\PHPExcel_Cell $cell, $name, $value)
    {
        preg_match_all('/([a-zA-Z]+)\((\w+)\)/', $name, $matches);
        for ($i = 0; $i < count($matches[0]); $i++) {
            $func = $matches[1][$i];
            $data = $matches[2][$i];
            if (array_key_exists($data, $this->dataCells)) {
                $value = sprintf('=%1$s(%2$s%3$d:%2$s%4$d)', $func, $this->dataCells[$data]['column'], $this->dataCells[$data]['rowStart'], $this->dataCells[$data]['rowEnd']);
            }
        }

        return $value;
    }

    /**
     * Fill range with value from the callback.
     * Every cells in range is checked if
     * it contain field sign.
     *
     * @param \PHPExcel_Worksheet $sheet  The worksheet
     * @param array $ranges  The data ranges
     * @return \NTLAB\Report\Util\Excel\Filer
     */
    protected function fillRange(\PHPExcel_Worksheet $sheet, $ranges = array(), $callback = null, $adjustRow = false)
    {
        // callback parameters (cell, data, value)
        if ($callback && is_callable($callback)) {
            $rangeStart = $ranges[0];
            $rangeEnd = $ranges[1];
            $cols = $rangeEnd[0] - $rangeStart[0];
            $rows = $rangeEnd[1] - $rangeStart[1];
            for ($row = 0; $row <= $rows; $row++) {
                for ($col = 0; $col <= $cols; $col++) {
                    $colindex = $rangeStart[0] + $col - 1;
                    $rowindex = $rangeStart[1] + $row;
                    $cell = $sheet->getCellByColumnAndRow($colindex, $rowindex);
                    // cell has value and prefixed with field signature
                    if (($value = $cell->getValue()) && (0 === strpos($value, $this->getFieldSign()))) {
                        $data = substr($value, strlen($this->getFieldSign()));
                        $value = call_user_func($callback, $cell, $data, $value);
                        // convert value if it is an object
                        if (is_object($value)) {
                            $value = (string) $value;
                        }
                        // set cell value back
                        $cell->setValue($value);
                    }
                    // if we're at the end of column, we should adjust the row height
                    if ($adjustRow && $col == $cols) {
                        if ($this->autoFit) {
                            // TODO: autofit row
                        } elseif (null !== $this->rowHeight) {
                            $cell->getParent()
                                ->getRowDimension($cell->getRow())
                                ->setRowHeight($this->rowHeight);
                        }
                    }
                }
            }
        }

        return $this;
    }

    /**
     * Find and detect bands position from Worksheet.
     *
     * @param \PHPExcel_Worksheet $sheet  The target sheet
     * @return \NTLAB\Report\Util\Excel\Filer
     */
    protected function findBands($sheet)
    {
        if (empty($this->bands)) {
            $rows = $sheet->getHighestDataRow();
            $cols = \PHPExcel_Cell::columnIndexFromString($sheet->getHighestDataColumn());
            $fsignFound = false;
            $fsignStart = null;
            $fsignEnd = null;
            $row = 1;
            while ($row <= $rows) {
                // check if row contain field sign
                $has_sign = false;
                for ($col = 1; $col <= $cols; $col++) {
                    $cell = $sheet->getCellByColumnAndRow($col - 1, $row);
                    if (($value = $cell->getValue()) && (0 === strpos($value, $this->getFieldSign()))) {
                        $has_sign = true;
                        break;
                    }
                }
                // master data band found in the row
                if ($has_sign) {
                    $fsignFound = true;
                    if (null === $fsignStart) {
                        $fsignStart = $row;
                    }
                    $fsignEnd = $row;
                } else {
                    // stop if field sign has previously found
                    if ($fsignFound) {
                        break;
                    }
                }
                $row++;
            }
            // bands found
            if ($fsignFound) {
                // title band
                $this->bands[self::BAND_TITLE] = sprintf('A1:%s%d', \PHPExcel_Cell::stringFromColumnIndex($cols - 1), $fsignStart - 1);
                // master data band
                $this->bands[self::BAND_MASTER_DATA] = sprintf('A%d:%s%d', $fsignStart, \PHPExcel_Cell::stringFromColumnIndex($cols - 1), $fsignEnd);
                // summary band
                $this->bands[self::BAND_SUMMARY] = sprintf('A%d:%s%d', $fsignEnd + 1, \PHPExcel_Cell::stringFromColumnIndex($cols - 1), $rows);
            }
        }

        return $this;
    }

    /**
     * Build the template and fill with objects data.
     *
     * @param array $objects  The objects
     * @param string $writerClass  The writer class
     * @return string
     */
    public function build($objects, $writerClass = null)
    {
        $this->objects = $objects;
        if (($tplXls = \PHPExcel_IOFactory::load($this->template)) && ($insheet = $tplXls->getSheetByName($this->sheet))) {
            $this->getScript()
                ->setObjects($this->objects)
            ;
            $resXls = $this->prepareResult($insheet, $outsheet);
            $this->findBands($insheet);
            $this->processBands($insheet, $outsheet);

            return $this->createOutput($resXls, $writerClass);
        }
    }

    /**
     * Prepare result sheet and assign document properties.
     *
     * @param \PHPExcel_Worksheet $source  Template sheet
     * @param \PHPExcel_Worksheet $sheet  Output sheet
     * @return \PHPExcel
     */
    protected function prepareResult($source, &$sheet)
    {
        $excel = new \PHPExcel();
        $sheet = $excel->getActiveSheet();
        $this->copySheetProp($source, $sheet);

        return $excel;
    }

    /**
     * Process report bands.
     *
     * @param \PHPExcel_Worksheet $insheet  Input worksheet
     * @param \PHPExcel_Worksheet $outsheet  Output worksheet
     * @return \NTLAB\Report\Util\Excel\Filer
     */
    protected function processBands($insheet, $outsheet)
    {
        $anchor = null;
        foreach ($this->allBands as $band) {
            if (!isset($this->bands[$band])) {
                continue;
            }
            $bandData = $this->bands[$band];
            switch ($band) {
                case static::BAND_MASTER_DATA:
                    $this->dataCells = array();
                    $this->getScript()
                        ->each(function(Script $script, Filer $_this) use ($insheet, $outsheet, &$anchor, $bandData) {
                            $_this->buildMasterData($insheet, $outsheet, $bandData, $anchor);
                        })
                    ;
                    break;

                default:
                    if (null == $this->getScript()->getContext() && count($this->objects)) {
                        $this->getScript()->setContext($this->objects[0]);
                    }
                    $ranges = $this->copyRange($insheet, $outsheet, $bandData, $anchor);
                    // process summary band and figure out the aggregate functions like SUM, AVG
                    if ($band === static::BAND_SUMMARY) {
                        $this->fillRange($outsheet, $ranges, array($this,  'fillSummary'));
                    }
                    break;
            }
        }

        return $this;
    }

    /**
     * Build detail row.
     *
     * @param \PHPExcel_Worksheet $insheet  Input sheet
     * @param \PHPExcel_Worksheet $outsheet  Output sheet
     * @param string $band  Band range address
     * @param array $anchor  Anchor range
     * @return \NTLAB\Report\Util\Excel\Filer
     */
    public function buildMasterData($insheet, $outsheet, $band, &$anchor)
    {
        $ranges = $this->copyRange($insheet, $outsheet, $band, $anchor, false);
        $this->fillRange($outsheet, $ranges, array($this, 'fillData'), true);

        return $this;
    }

    /**
     * Create ouput content from result sheet.
     *
     * @param \PHPExcel $xls  Excel output
     * @param string $writerClass  Writer class
     * @return string
     */
    protected function createOutput($xls, $writerClass)
    {
        $writerClass = null !== $writerClass ? $writerClass : $this->defaultWriter;
        if (null == $writerClass) {
            switch ($readerClass = \PHPExcel_IOFactory::identify($this->template)) {
                case 'Excel2007':
                    $writerClass = $readerClass;
                    break;

                default:
                    $writerClass = 'Excel5';
                    break;
            }
        }
        $filename = tempnam(dirname($this->template), 'xlstmp');
        $writer = \PHPExcel_IOFactory::createWriter($xls, $writerClass);
        $writer->save($filename);

        return file_get_contents($filename);
    }
}