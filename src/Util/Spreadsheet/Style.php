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

use PhpOffice\PhpSpreadsheet\Style as XlStyle;
use PhpOffice\PhpSpreadsheet\Style\Alignment as XlStyleAlignment;
use PhpOffice\PhpSpreadsheet\Style\Border as XlStyleBorder;
use PhpOffice\PhpSpreadsheet\Style\Borders as XlStyleBorders;
use PhpOffice\PhpSpreadsheet\Style\Color as XlStyleColor;
use PhpOffice\PhpSpreadsheet\Style\Fill as XlStyleFill;
use PhpOffice\PhpSpreadsheet\Style\Font as XlStyleFont;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat as XlStyleNumberFormat;
use PhpOffice\PhpSpreadsheet\Style\Protection as XlStyleProtection;

class Style
{
    protected static $properties = array(
        'PhpOffice\PhpSpreadsheet\Style' => array(
            'fill' => 'getFill',
            'font' => 'getFont',
            'borders' => 'getBorders',
            'alignment' => 'getAlignment',
            'numberformat' => 'getNumberFormat',
            'protection' => 'getProtection'
        ),
        'PhpOffice\PhpSpreadsheet\Style\Fill' => array(
            'type' => 'getFillType',
            'rotation' => 'getRotation',
            'startcolor' => 'getStartColor',
            'endcolor' => 'getEndColor'
        ),
        'PhpOffice\PhpSpreadsheet\Style\Font' => array(
            'name' => 'getName',
            'bold' => 'getBold',
            'italic' => 'getItalic',
            'superScript' => array('getSuperScript',  true),
            'subScript' => array('getSubScript', true),
            'underline' => 'getUnderline',
            'strike' => 'getStrikeThrough',
            'color' => 'getColor',
            'size' => 'getSize'
        ),
        'PhpOffice\PhpSpreadsheet\Style\Borders' => array(
            'left' => 'getLeft',
            'right' => 'getRight',
            'top' => 'getTop',
            'bottom' => 'getBottom',
            'diagonal' => 'getDiagonal',
            'diagonaldirection' => 'getDiagonalDirection'
            // 'allborders' => 'getAllBorders',
        ),
        'PhpOffice\PhpSpreadsheet\Style\Alignment' => array(
            'horizontal' => 'getHorizontal',
            'vertical' => 'getVertical',
            'rotation' => 'getTextRotation',
            'wrap' => 'getWrapText',
            'shrinkToFit' => 'getShrinkToFit',
            'indent' => 'getIndent'
        ),
        'PhpOffice\PhpSpreadsheet\Style\NumberFormat' => array(
            'code' => 'getFormatCode'
        ),
        'PhpOffice\PhpSpreadsheet\Style\Protection' => array(
            'locked' => 'getLocked',
            'hidden' => 'getHidden'
        ),
        'PhpOffice\PhpSpreadsheet\Style\Color' => array(
            'argb' => 'getARGB'
        ),
        'PhpOffice\PhpSpreadsheet\Style\Border' => array(
            'style' => 'getBorderStyle',
            'color' => 'getColor'
        )
    );

    /**
     * Convert style object to array.
     *
     * @param mixed $object  The object
     * @return mixed
     */
    protected static function objectToArray($object)
    {
        if (is_object($object)) {
            $class = get_class($object);
            if (isset(self::$properties[$class])) {
                $result = array();
                foreach (self::$properties[$class] as $prop => $method) {
                    $mutual = false;
                    if (is_array($method)) {
                        $mutual = $method[1];
                        $method = $method[0];
                    }
                    $value = $object->$method();
                    if ($mutual && !$value) {
                        continue;
                    }
                    $result[$prop] = self::objectToArray($value);
                }

                return $result;
            } else {
                return clone $object;
            }
        } else {
            return $object;
        }
    }

    /**
     * Get style array from style object.
     *
     * @param XlStyle $style  The style
     * @return array
     */
    public static function styleToArray(XlStyle $style)
    {
        return self::objectToArray($style);
    }

    /**
     * Copy alignment.
     *
     * @param XlStyleAlignment $source  The source
     * @param XlStyleAlignment $dest  The destination
     */
    public static function copyAlignment(XlStyleAlignment $source, XlStyleAlignment $dest)
    {
        if ($dest->getHashCode() !== $source->getHashCode()) {
            if ($dest->getHorizontal() !== $source->getHorizontal()) {
                $dest->setHorizontal($source->getHorizontal());
            }
            if ($dest->getIndent() !== $source->getIndent()) {
                $dest->setIndent($source->getIndent());
            }
            if ($dest->getShrinkToFit() !== $source->getShrinkToFit()) {
                $dest->setShrinkToFit($source->getShrinkToFit());
            }
            if ($dest->getTextRotation() !== $source->getTextRotation()) {
                $dest->setTextRotation($source->getTextRotation());
            }
            if ($dest->getVertical() !== $source->getVertical()) {
                $dest->setVertical($source->getVertical());
            }
            if ($dest->getWrapText() !== $source->getWrapText()) {
                $dest->setWrapText($source->getWrapText());
            }
        }
    }

    /**
     * Copy color.
     *
     * @param XlStyleColor $source  The source
     * @param XlStyleColor $dest  The destination
     */
    public static function copyColor(XlStyleColor $source, XlStyleColor $dest)
    {
        if ($dest->getHashCode() !== $source->getHashCode()) {
            $dest->setARGB($source->getARGB());
        }
    }

    /**
     * Copy border.
     *
     * @param XlStyleBorder $source  The source
     * @param XlStyleBorder $dest  The destination
     */
    public static function copyBorder(XlStyleBorder $source, XlStyleBorder $dest)
    {
        if ($dest->getHashCode() !== $source->getHashCode()) {
            if ($dest->getBorderStyle() !== $source->getBorderStyle()) {
                $dest->setBorderStyle($source->getBorderStyle());
            }
            self::copyColor($source->getColor(), $dest->getColor());
        }
    }

    /**
     * Copy borders.
     *
     * @param XlStyleBorders $source  The source
     * @param XlStyleBorders $dest  The destination
     */
    public static function copyBorders(XlStyleBorders $source, XlStyleBorders $dest)
    {
        if ($dest->getHashCode() !== $source->getHashCode()) {
            self::copyBorder($source->getLeft(), $dest->getLeft());
            self::copyBorder($source->getRight(), $dest->getRight());
            self::copyBorder($source->getTop(), $dest->getTop());
            self::copyBorder($source->getBottom(), $dest->getBottom());
            self::copyBorder($source->getDiagonal(), $dest->getDiagonal());
            if ($dest->getDiagonalDirection() !== $source->getDiagonalDirection()) {
                $dest->setDiagonalDirection($source->getDiagonalDirection());
            }
            /*
             * self::copyBorder($source->getAllBorders(), $dest->getAllBorders());
             * self::copyBorder($source->getOutline(), $dest->getOutline());
             * self::copyBorder($source->getInside(), $dest->getInside());
             * self::copyBorder($source->getVertical(), $dest->getVertical());
             * self::copyBorder($source->getHorizontal(), $dest->getHorizontal());
             */
        }
    }

    /**
     * Copy fill.
     *
     * @param XlStyleFill $source  The source
     * @param XlStyleFill $dest  The destination
     */
    public static function copyFill(XlStyleFill $source, XlStyleFill $dest)
    {
        if ($dest->getHashCode() !== $source->getHashCode()) {
            if ($dest->getFillType() !== $source->getFillType()) {
                $dest->setFillType($source->getFillType());
            }
            if ($dest->getRotation() !== $source->getRotation()) {
                $dest->setRotation($source->getRotation());
            }
            self::copyColor($source->getStartColor(), $dest->getStartColor());
            self::copyColor($source->getEndColor(), $dest->getEndColor());
        }
    }

    /**
     * Copy font.
     *
     * @param XlStyleFont $source  The source
     * @param XlStyleFont $dest  The destination
     */
    public static function copyFont(XlStyleFont $source, XlStyleFont $dest)
    {
        if ($dest->getHashCode() !== $source->getHashCode()) {
            if ($dest->getName() !== $source->getName()) {
                $dest->setName($source->getName());
            }
            if ($dest->getSize() !== $source->getSize()) {
                $dest->setSize($source->getSize());
            }
            self::copyColor($source->getColor(), $dest->getColor());
            if ($dest->getBold() !== $source->getBold()) {
                $dest->setBold($source->getBold());
            }
            if ($dest->getItalic() !== $source->getItalic()) {
                $dest->setItalic($source->getItalic());
            }
            if ($dest->getUnderline() !== $source->getUnderline()) {
                $dest->setUnderline($source->getUnderline());
            }
            if ($dest->getStrikethrough() !== $source->getStrikethrough()) {
                $dest->setStrikethrough($source->getStrikethrough());
            }
            if ($source->getSubScript()) {
                $dest->setSubScript($source->getSubScript());
            }
            if ($source->getSuperScript()) {
                $dest->setSuperScript($source->getSuperScript());
            }
        }
    }

    /**
     * Copy number format.
     *
     * @param XlStyleNumberFormat $source  The source
     * @param XlStyleNumberFormat $dest  The destination
     */
    public static function copyNumberFormat(XlStyleNumberFormat $source, XlStyleNumberFormat $dest)
    {
        if ($dest->getHashCode() !== $source->getHashCode()) {
            if ($dest->getBuiltInFormatCode() !== $source->getBuiltInFormatCode()) {
                $dest->setBuiltInFormatCode($source->getBuiltInFormatCode());
            }
            if ($dest->getFormatCode() !== $source->getFormatCode()) {
                $dest->setFormatCode($source->getFormatCode());
            }
        }
    }

    /**
     * Copy cell style.
     *
     * @param XlStyle $source  The source
     * @param XlStyle $dest  The destination
     */
    public static function copyStyle(XlStyle $source, XlStyle $dest)
    {
        if ($dest->getHashCode() !== $source->getHashCode()) {
            self::copyFill($source->getFill(), $dest->getFill());
            self::copyFont($source->getFont(), $dest->getFont());
            self::copyBorders($source->getBorders(), $dest->getBorders());
            self::copyAlignment($source->getAlignment(), $dest->getAlignment());
            self::copyNumberFormat($source->getNumberFormat(), $dest->getNumberFormat());
            // $source->getConditionalStyles();
            // $source->getProtection();
        }
    }
}