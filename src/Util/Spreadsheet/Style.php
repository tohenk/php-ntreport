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

namespace NTLAB\Report\Util\Spreadsheet;

use PhpOffice\PhpSpreadsheet\Exception as XlException;
use PhpOffice\PhpSpreadsheet\Style\Alignment as XlStyleAlignment;
use PhpOffice\PhpSpreadsheet\Style\Border as XlStyleBorder;
use PhpOffice\PhpSpreadsheet\Style\Borders as XlStyleBorders;
use PhpOffice\PhpSpreadsheet\Style\Color as XlStyleColor;
use PhpOffice\PhpSpreadsheet\Style\Fill as XlStyleFill;
use PhpOffice\PhpSpreadsheet\Style\Font as XlStyleFont;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat as XlStyleNumberFormat;
use PhpOffice\PhpSpreadsheet\Style\Protection as XlStyleProtection;
use PhpOffice\PhpSpreadsheet\Style\Style as XlStyle;

class Style
{
    protected static $properties = array(
        'PhpOffice\PhpSpreadsheet\Style\Alignment' => array(
            'horizontal',
            'vertical',
            'textRotation',
            'wrapText',
            'shrinkToFit',
            'indent',
            'readOrder',
        ),
        'PhpOffice\PhpSpreadsheet\Style\Border' => array(
            'borderStyle',
            'color',
        ),
        'PhpOffice\PhpSpreadsheet\Style\Borders' => array(
            'left',
            'right',
            'top',
            'bottom',
            'diagonal',
            'diagonalDirection',
            'allBorders',
        ),
        'PhpOffice\PhpSpreadsheet\Style\Color' => array(
            'argb',
        ),
        'PhpOffice\PhpSpreadsheet\Style\Fill' => array(
            'fillType',
            'rotation',
            'startColor',
            'endColor',
        ),
        'PhpOffice\PhpSpreadsheet\Style\Font' => array(
            'name',
            'bold',
            'italic',
            'superscript',
            'subscript',
            'underline',
            'strikethrough',
            'color',
            'size',
        ),
        'PhpOffice\PhpSpreadsheet\Style\NumberFormat' => array(
            'formatCode',
        ),
        'PhpOffice\PhpSpreadsheet\Style\Protection' => array(
            'locked',
            'hidden',
        ),
        'PhpOffice\PhpSpreadsheet\Style\Style' => array(
            'fill',
            'font',
            'borders',
            'alignment',
            'numberFormat',
            'protection',
            'quotePrefix',
        ),
    );

    /**
     * Get style array from style object.
     *
     * @param mixed $value  The style
     * @return array
     */
    public static function styleToArray($value)
    {
        if (is_object($value)) {
            $class = get_class($value);
            if (isset(self::$properties[$class])) {
                $result = array();
                foreach (self::$properties[$class] as $key) {
                    $method = sprintf('get%s', ucfirst($key));
                    try {
                        $prop = $value->$method();
                        // set property only on boolean true
                        if (is_bool($prop) && false === $prop) {
                            continue;
                        }
                        $result[$key] = self::styleToArray($prop);
                    } catch (XlException $e) {
                        // ignore error
                    }
                }
                return $result;
            } else {
                return clone $value;
            }
        } else {
            return $value;
        }
    }
}