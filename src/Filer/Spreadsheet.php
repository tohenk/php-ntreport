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

namespace NTLAB\Report\Filer;

use NTLAB\Report\Session\Session;
use NTLAB\Report\Util\Spreadsheet\Style;
use NTLAB\Report\Util\Spreadsheet\RichText;
use NTLAB\Script\Context\PartialObject;
use NTLAB\Script\Core\Script;
use PhpOffice\PhpSpreadsheet\Cell\Cell as XlCell;
use PhpOffice\PhpSpreadsheet\Cell\CellAddress as XlCellAddress;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate as XlCoordinate;
use PhpOffice\PhpSpreadsheet\IOFactory as XlIOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet as XlSpreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet as XlWorksheet;

/**
 * Spreadsheet filer using PHPOffice PhpSpreadsheet.
 *
 * @author Toha <tohenk@yahoo.com>
 */
class Spreadsheet implements FilerInterface
{
    use Filer;

    public const BAND_TITLE = 'title';
    public const BAND_HEADER = 'header';
    public const BAND_MASTER_DATA = 'master-data';
    public const BAND_FOOTER = 'footer';
    public const BAND_SUMMARY = 'summary';

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
     * Copy worksheet property.
     *
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $source  Source worksheet
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $dest  Destination worksheet
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
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $sheet  The worksheet object
     * @param string $cell  The cell
     * @return string
     */
    protected function getCellMerged(XlWorksheet $sheet, $cell)
    {
        foreach ($sheet->getMergeCells() as $key => $range) {
            list($rangeStart, ) = explode(':', $key);
            if ($rangeStart === $cell) {
                return $range;
            }
        }
    }

    /**
     * Merge cells.
     *
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $sheet  The worksheet
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
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $source  The source worksheet
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $dest  The destination worksheet
     * @param string $range  The source range
     * @param array $anchor  The destination anchor
     * @param bool $replaceTag  True to replace tag before applying value to destination sheet
     * @return array Range boundaries array(start, end)
     */
    protected function copyRange(XlWorksheet $source, XlWorksheet $dest, $range, &$anchor, $replaceTag = true)
    {
        list($rangeStart, $rangeEnd) = XlCoordinate::rangeBoundaries($range);
        if (null === $anchor) {
            $anchor = $rangeStart;
            // apply offset
            if (($offset = $this->getScript()->getIterator()->getStart()) > 0) {
                $anchor[1] += $offset;
            }
        }
        $cols = $rangeEnd[0] - $rangeStart[0];
        $rows = $rangeEnd[1] - $rangeStart[1];
        for ($row = 0; $row <= $rows; $row++) {
            for ($col = 0; $col <= $cols; $col++) {
                $scell = $source->getCell(XlCellAddress::fromColumnAndRow($rangeStart[0] + $col, $rangeStart[1] + $row));
                $dcell = $dest->getCell(XlCellAddress::fromColumnAndRow($anchor[0] + $col, $anchor[1] + $row));
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
                if ($col === 0) {
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
        $anchor[1] += $rows + 1;

        return $ranges;
    }

    /**
     * Fill range for master-data.
     *
     * @param \PhpOffice\PhpSpreadsheet\Cell\Cell $cell  The cell
     * @param string $name  Data name
     * @param string $value  Cell value
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
     * @param \PhpOffice\PhpSpreadsheet\Cell\Cell $cell  The cell
     * @param string $name  Data name
     * @param string $value  Cell value
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
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $sheet  The worksheet
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
                    $cell = $sheet->getCell(XlCellAddress::fromColumnAndRow($colindex, $rowindex));
                    // cell has value and prefixed with field signature
                    if (($value = $cell->getValue()) && (0 === strpos($value, $this->getFieldSign()))) {
                        $data = substr($value, strlen($this->getFieldSign()));
                        $value = call_user_func($callback, $cell, $data, $value);
                        // convert date time
                        if ($value instanceof \DateTime) {
                            $value = $value->format(\DateTime::ISO8601);
                        }
                        // convert value if it is an object
                        if (is_object($value)) {
                            $value = (string) $value;
                        }
                        // set cell value back
                        $cell->setValue($value);
                    }
                    // if we're at the end of column, we should adjust the row height
                    if ($adjustRow && $col === $cols) {
                        if ($this->autoFit) {
                            // TODO: autofit row
                        } elseif (null !== $this->rowHeight) {
                            $cell->getWorksheet()
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
     * Find and detect bands position from work sheet.
     *
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $sheet  The target sheet
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
                    $cell = $sheet->getCell(XlCellAddress::fromColumnAndRow($col, $row));
                    // ignore aggregate sign
                    if (($value = $cell->getValue()) && (0 === strpos($value, $this->getFieldSign())) && $this->aggregateSign !== substr($value, 1, 1)) {
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
        if (($tpl = XlIOFactory::load($this->template)) && ($insheet = $tpl->getSheetByName($this->sheet))) {
            $this->getScript()
                ->setObjects($this->objects)
            ;
            $outsheet = null;
            if (($output = $this->session->read(Session::OUT)) && is_readable($output)) {
                $res = XlIOFactory::load($output);
                $outsheet = $res->getActiveSheet();
            } else {
                $res = $this->prepareResult($insheet, $outsheet);
            }
            $this->findBands($insheet);
            $this->processBands($insheet, $outsheet);

            return $this->createOutput($res, $writerClass);
        }
    }

    /**
     * Prepare result sheet and assign document properties.
     *
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $source  Template sheet
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $sheet  Output sheet
     * @return \PhpOffice\PhpSpreadsheet\Spreadsheet
     */
    protected function prepareResult($source, &$sheet)
    {
        $xls = new XlSpreadsheet();
        $sheet = $xls->getActiveSheet();
        $this->copySheetProp($source, $sheet);

        return $xls;
    }

    /**
     * Process report bands.
     *
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $insheet  Input worksheet
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $outsheet  Output worksheet
     * @return \NTLAB\Report\Filer\Spreadsheet
     */
    protected function processBands($insheet, $outsheet)
    {
        $anchor = null;
        $iterator = $this->getScript()->getIterator();
        foreach ($this->allBands as $band) {
            if (!isset($this->bands[$band])) {
                continue;
            }
            $bandData = $this->bands[$band];
            switch ($band) {
                case static::BAND_MASTER_DATA:
                    $this->dataCells = [];
                    $this->getScript()
                        ->each(function (Script $script, Spreadsheet $_this) use ($insheet, $outsheet, &$anchor, $bandData) {
                            $_this->buildMasterData($insheet, $outsheet, $bandData, $anchor);
                        })
                    ;
                    break;
                default:
                    $first = $iterator->getStart() === 0;
                    $last = $iterator->getRecNo() === $iterator->getRecCount();
                    if ($band === static::BAND_TITLE && !$first) {
                        break;
                    }
                    if ($band === static::BAND_SUMMARY && !$last) {
                        break;
                    }
                    $objects = $this->objects instanceof PartialObject ? $this->objects->getObjects() :
                        $this->objects;
                    if (null === $this->getScript()->getContext() && count($objects)) {
                        $this->getScript()->setContext($objects[0]);
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
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $insheet  Input sheet
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $outsheet  Output sheet
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
     * @param \PhpOffice\PhpSpreadsheet\Spreadsheet $xls  Excel output
     * @param string $writerClass  Writer class
     * @return string
     */
    protected function createOutput($xls, $writerClass)
    {
        $writerClass = null !== $writerClass ? $writerClass : $this->defaultWriter;
        if (null === $writerClass) {
            switch ($readerClass = XlIOFactory::identify($this->template)) {
                case 'Xlsx':
                    $writerClass = $readerClass;
                    break;
                default:
                    $writerClass = 'Xls';
                    break;
            }
        }
        if ('~' === substr($filename = basename($this->template), 0, 1)) {
            $filename = substr($filename, 1);
        }
        $filename = $this->session->createWorkDir().DIRECTORY_SEPARATOR.$filename;
        $writer = XlIOFactory::createWriter($xls, $writerClass);
        $writer->save($filename);
        $this->session->store(Session::OUT, $filename);

        return file_get_contents($filename);
    }
}
