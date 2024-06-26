<?php

/*
 * The MIT License
 *
 * Copyright (c) 2014-2024 Toha <tohenk@yahoo.com>
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

namespace NTLAB\Report\Filer;

use NTLAB\Script\Core\Script;
use PhpOffice\PhpSpreadsheet\IOFactory as XlIOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet as XlSpreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet as XlWorksheet;
use PhpOffice\PhpSpreadsheet\Cell\Cell as XlCell;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate as XlCoordinate;
use NTLAB\Report\Util\Spreadsheet\Style;
use NTLAB\Report\Util\Spreadsheet\RichText;

/**
 * Spreadsheet filer using PHPOffice PhpSpreadsheet.
 *
 * @author Toha
 */
class Spreadsheet implements FilerInterface
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
    protected $bands = [];

    /**
     * @var array
     */
    protected $datas = [];

    /**
     * @var array
     */
    protected $allBands = [
        self::BAND_TITLE,
        self::BAND_HEADER,
        self::BAND_MASTER_DATA,
        self::BAND_FOOTER,
        self::BAND_SUMMARY
    ];

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
     * @var string
     */
    protected $aggregateSign = '=';

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
    protected $dataCells = [];

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
     * @return \NTLAB\Report\Filer\Spreadsheet
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
     * @return \NTLAB\Report\Filer\Spreadsheet
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
     * @return \NTLAB\Report\Filer\Spreadsheet
     */
    public function setFieldSign($signature)
    {
        $this->fieldSign = $signature;
        return $this;
    }

    /**
     * Get aggregate signature.
     *
     * @return string
     */
    public function getAggregateSign()
    {
        return $this->aggregateSign;
    }

    /**
     * Set aggregate signature.
     *
     * @param string $signature  The signature
     * @return \NTLAB\Report\Filer\Spreadsheet
     */
    public function setAggregateSign($signature)
    {
        $this->aggregateSign = $signature;
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
     * @return \NTLAB\Report\Filer\Spreadsheet
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
     * @return \NTLAB\Report\Filer\Spreadsheet
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
     * @return \NTLAB\Report\Filer\Spreadsheet
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
     * @return \NTLAB\Report\Filer\Spreadsheet
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
     * @return \NTLAB\Report\Filer\Spreadsheet
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
        if (null === $this->script) {
            $this->script = new Script();
        }
        return $this->script;
    }

    /**
     * Set the script object.
     *
     * @param \NTLAB\Script\Core\Script $script  The script object
     * @return \NTLAB\Report\Filer\Spreadsheet
     */
    public function setScript(Script $script)
    {
        $this->script = $script;
        return $this;
    }

    /**
     * Copy worksheet property.
     *
     * @param XlWorksheet $source  Source worksheet
     * @param XlWorksheet $dest  Destination worksheet
     * @return \NTLAB\Report\Filer\Spreadsheet
     */
    protected function copySheetProp(XlWorksheet $source, XlWorksheet $dest)
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
     * @param XlWorksheet $sheet  The worksheet object
     * @param string $cell  The cell
     * @return string
     */
    protected function getCellMerged(XlWorksheet $sheet, $cell)
    {
        foreach ($sheet->getMergeCells() as $key => $range) {
            list($rangeStart, ) = explode(':', $key);
            if ($rangeStart == $cell) {
                return $range;
            }
        }
    }

    /**
     * Merge cells.
     *
     * @param XlWorksheet $sheet  The worksheet
     * @param string $cell  The origin cell
     * @param int $width  The number of columns from origin to merge
     * @param int $height  The number of rows from origin to merge
     * @return \NTLAB\Report\Filer\Spreadsheet
     */
    protected function mergeCells(XlWorksheet $sheet, $cell, $width, $height)
    {
        list($rangeStart, $rangeEnd) = XlCoordinate::rangeBoundaries($cell);
        $rangeEnd[0] = (int) $rangeStart[0] + $width - 1;
        $rangeEnd[1] = (int) $rangeStart[1] + $height - 1;
        $range = XlCoordinate::stringFromColumnIndex($rangeStart[0]).$rangeStart[1].':'.XlCoordinate::stringFromColumnIndex($rangeEnd[0]).$rangeEnd[1];
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
        $matches = null;
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
     * @param XlWorksheet $source  The source worksheet
     * @param XlWorksheet $dest  The destination worksheet
     * @param string $range  The source range
     * @param array $anchor  The destination anchor
     * @param bool $replaceTag  True to replace tag before applying value to destination sheet
     * @return array Range boundaries array(start, end)
     */
    protected function copyRange(XlWorksheet $source, XlWorksheet $dest, $range, &$anchor, $replaceTag = true)
    {
        list($rangeStart, $rangeEnd) = XlCoordinate::rangeBoundaries($range);
        if (null == $anchor) {
            $anchor = $rangeStart;
        }
        $cols = $rangeEnd[0] - $rangeStart[0];
        $rows = $rangeEnd[1] - $rangeStart[1];
        for ($row = 0; $row <= $rows; $row++) {
            for ($col = 0; $col <= $cols; $col++) {
                $scell = $source->getCellByColumnAndRow($rangeStart[0] + $col, $rangeStart[1] + $row);
                $dcell = $dest->getCellByColumnAndRow($anchor[0] + $col, $anchor[1] + $row);
                // copy value, ignore empty cell
                if ($svalue = $scell->getValue()) {
                    $dcell->setValue($replaceTag ? $this->replaceTag($svalue) : $svalue);
                }
                // copy style
                $dest->getStyle($dcell->getCoordinate())
                    ->applyFromArray(Style::styleToArray($source->getStyle($scell->getCoordinate())));
                // merge cells
                if ($merged = $this->getCellMerged($source, $scell->getCoordinate())) {
                    $dim = XlCoordinate::rangeDimension($merged);
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
        $ranges = [$anchor, [$anchor[0] + $cols, $anchor[1] + $rows]];
        // increment anchor rows
        $anchor[1] = $anchor[1] + $rows + 1;
        return $ranges;
    }

    /**
     * Fill range for master-data.
     *
     * @param XlCell $cell  The cell
     * @param string $name  Data name
     * @param string $value  XlCell value
     * @return mixed
     */
    protected function fillData(XlCell $cell, $name, $value)
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
            $this->dataCells[$name] = [
                'column' => $cell->getColumn(),
                'rowStart' => $cell->getRow(),
                'rowEnd' => $cell->getRow()
            ];
        } else {
            $this->dataCells[$name]['rowEnd'] = $cell->getRow();
        }
        return $value;
    }

    /**
     * Fill range for summary.
     *
     * @param XlCell $cell  The cell
     * @param string $name  Data name
     * @param string $value  XlCell value
     * @return mixed
     */
    protected function fillSummary(XlCell $cell, $name, $value)
    {
        $value = $name;
        $matches = null;
        preg_match_all('/([a-zA-Z]+)\(([a-zA-Z0-9_]+)\)/', $name, $matches);
        for ($i = 0; $i < count($matches[0]); $i++) {
            $func = $matches[1][$i];
            $data = $matches[2][$i];
            if (in_array(strtoupper($func), ['SUM', 'AVG', 'AVERAGE']) && array_key_exists($data, $this->dataCells)) {
                $value = str_replace($matches[0][$i], sprintf('%1$s(%2$s%3$d:%2$s%4$d)', $func, $this->dataCells[$data]['column'], $this->dataCells[$data]['rowStart'], $this->dataCells[$data]['rowEnd']), $value);
            }
        }
        return $value;
    }

    /**
     * Fill range with value from the callback.
     * Every cells in range is checked if
     * it contain field sign.
     *
     * @param XlWorksheet $sheet  The worksheet
     * @param array $ranges  The data ranges
     * @return \NTLAB\Report\Filer\Spreadsheet
     */
    protected function fillRange(XlWorksheet $sheet, $ranges = [], $callback = null, $adjustRow = false)
    {
        // callback parameters (cell, data, value)
        if ($callback && is_callable($callback)) {
            $rangeStart = $ranges[0];
            $rangeEnd = $ranges[1];
            $cols = $rangeEnd[0] - $rangeStart[0];
            $rows = $rangeEnd[1] - $rangeStart[1];
            for ($row = 0; $row <= $rows; $row++) {
                for ($col = 0; $col <= $cols; $col++) {
                    $colindex = $rangeStart[0] + $col;
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
     * Find and detect bands position from XlWorksheet.
     *
     * @param XlWorksheet $sheet  The target sheet
     * @return \NTLAB\Report\Filer\Spreadsheet
     */
    protected function findBands($sheet)
    {
        if (empty($this->bands)) {
            $rows = $sheet->getHighestDataRow();
            $cols = XlCoordinate::columnIndexFromString($sheet->getHighestDataColumn());
            $fsignFound = false;
            $fsignStart = null;
            $fsignEnd = null;
            $row = 1;
            while ($row <= $rows) {
                // check if row contain field sign
                $has_sign = false;
                for ($col = 1; $col <= $cols; $col++) {
                    $cell = $sheet->getCellByColumnAndRow($col, $row);
                    // ignore aggregate sign
                    if (($value = $cell->getValue()) && (0 === strpos($value, $this->getFieldSign())) && $this->aggregateSign != substr($value, 1, 1)) {
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
                $this->bands[self::BAND_TITLE] = sprintf('A1:%s%d', XlCoordinate::stringFromColumnIndex($cols), $fsignStart - 1);
                // master data band
                $this->bands[self::BAND_MASTER_DATA] = sprintf('A%d:%s%d', $fsignStart, XlCoordinate::stringFromColumnIndex($cols), $fsignEnd);
                // summary band
                $this->bands[self::BAND_SUMMARY] = sprintf('A%d:%s%d', $fsignEnd + 1, XlCoordinate::stringFromColumnIndex($cols), $rows);
            }
        }
        return $this;
    }

    /**
     * {@inheritDoc}
     * @see \NTLAB\Report\Filer\FilerInterface::build()
     */
    public function build($template, $objects, $writerClass = null)
    {
        $this->template = $template;
        $this->objects = $objects;
        if (($tplXls = XlIOFactory::load($this->template)) && ($insheet = $tplXls->getSheetByName($this->sheet))) {
            $this->getScript()
                ->setObjects($this->objects)
            ;
            $outsheet = null;
            $resXls = $this->prepareResult($insheet, $outsheet);
            $this->findBands($insheet);
            $this->processBands($insheet, $outsheet);
            return $this->createOutput($resXls, $writerClass);
        }
    }

    /**
     * Prepare result sheet and assign document properties.
     *
     * @param XlWorksheet $source  Template sheet
     * @param XlWorksheet $sheet  Output sheet
     * @return XlSpreadsheet
     */
    protected function prepareResult($source, &$sheet)
    {
        $excel = new XlSpreadsheet();
        $sheet = $excel->getActiveSheet();
        $this->copySheetProp($source, $sheet);
        return $excel;
    }

    /**
     * Process report bands.
     *
     * @param XlWorksheet $insheet  Input worksheet
     * @param XlWorksheet $outsheet  Output worksheet
     * @return \NTLAB\Report\Filer\Spreadsheet
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
                    $this->dataCells = [];
                    $this->getScript()
                        ->each(function(Script $script, Spreadsheet $_this) use ($insheet, $outsheet, &$anchor, $bandData) {
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
                        $this->fillRange($outsheet, $ranges, [$this,  'fillSummary']);
                    }
                    break;
            }
        }
        return $this;
    }

    /**
     * Build detail row.
     *
     * @param XlWorksheet $insheet  Input sheet
     * @param XlWorksheet $outsheet  Output sheet
     * @param string $band  Band range address
     * @param array $anchor  Anchor range
     * @return \NTLAB\Report\Filer\Spreadsheet
     */
    public function buildMasterData($insheet, $outsheet, $band, &$anchor)
    {
        $ranges = $this->copyRange($insheet, $outsheet, $band, $anchor, false);
        $this->fillRange($outsheet, $ranges, [$this, 'fillData'], true);
        return $this;
    }

    /**
     * Create ouput content from result sheet.
     *
     * @param XlSpreadsheet $xls  Excel output
     * @param string $writerClass  Writer class
     * @return string
     */
    protected function createOutput($xls, $writerClass)
    {
        $writerClass = null !== $writerClass ? $writerClass : $this->defaultWriter;
        if (null == $writerClass) {
            switch ($readerClass = XlIOFactory::identify($this->template)) {
                case 'Xlsx':
                    $writerClass = $readerClass;
                    break;
                default:
                    $writerClass = 'Xls';
                    break;
            }
        }
        $filename = tempnam(dirname($this->template), 'xls');
        $writer = XlIOFactory::createWriter($xls, $writerClass);
        $writer->save($filename);
        return file_get_contents($filename);
    }
}