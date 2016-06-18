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

class Style
{
    protected static $properties = array(
        'PHPExcel_Style' => array(
            'fill' => 'getFill',
            'font' => 'getFont',
            'borders' => 'getBorders',
            'alignment' => 'getAlignment',
            'numberformat' => 'getNumberFormat',
            'protection' => 'getProtection'
        ),
        'PHPExcel_Style_Fill' => array(
            'type' => 'getFillType',
            'rotation' => 'getRotation',
            'startcolor' => 'getStartColor',
            'endcolor' => 'getEndColor'
        ),
        'PHPExcel_Style_Font' => array(
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
        'PHPExcel_Style_Borders' => array(
            'left' => 'getLeft',
            'right' => 'getRight',
            'top' => 'getTop',
            'bottom' => 'getBottom',
            'diagonal' => 'getDiagonal',
            'diagonaldirection' => 'getDiagonalDirection'
            // 'allborders' => 'getAllBorders',
        ),
        'PHPExcel_Style_Alignment' => array(
            'horizontal' => 'getHorizontal',
            'vertical' => 'getVertical',
            'rotation' => 'getTextRotation',
            'wrap' => 'getWrapText',
            'shrinkToFit' => 'getShrinkToFit',
            'indent' => 'getIndent'
        ),
        'PHPExcel_Style_NumberFormat' => array(
            'code' => 'getFormatCode'
        ),
        'PHPExcel_Style_Protection' => array(
            'locked' => 'getLocked',
            'hidden' => 'getHidden'
        ),
        'PHPExcel_Style_Color' => array(
            'argb' => 'getARGB'
        ),
        'PHPExcel_Style_Border' => array(
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
     * @param \PHPExcel_Style $style  The style
     * @return array
     */
    public static function styleToArray(\PHPExcel_Style $style)
    {
        return self::objectToArray($style);
    }

    /**
     * Copy alignment.
     *
     * @param \PHPExcel_Style_Alignment $source  The source
     * @param \PHPExcel_Style_Alignment $dest  The destination
     */
    public static function copyAlignment(\PHPExcel_Style_Alignment $source, \PHPExcel_Style_Alignment $dest)
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
     * @param \PHPExcel_Style_Color $source  The source
     * @param \PHPExcel_Style_Color $dest  The destination
     */
    public static function copyColor(\PHPExcel_Style_Color $source, \PHPExcel_Style_Color $dest)
    {
        if ($dest->getHashCode() !== $source->getHashCode()) {
            $dest->setARGB($source->getARGB());
        }
    }

    /**
     * Copy border.
     *
     * @param \PHPExcel_Style_Border $source  The source
     * @param \PHPExcel_Style_Border $dest  The destination
     */
    public static function copyBorder(\PHPExcel_Style_Border $source, \PHPExcel_Style_Border $dest)
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
     * @param \PHPExcel_Style_Borders $source  The source
     * @param \PHPExcel_Style_Borders $dest  The destination
     */
    public static function copyBorders(\PHPExcel_Style_Borders $source, \PHPExcel_Style_Borders $dest)
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
     * @param \PHPExcel_Style_Fill $source  The source
     * @param \PHPExcel_Style_Fill $dest  The destination
     */
    public static function copyFill(\PHPExcel_Style_Fill $source, \PHPExcel_Style_Fill $dest)
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
     * @param \PHPExcel_Style_Font $source  The source
     * @param \PHPExcel_Style_Font $dest  The destination
     */
    public static function copyFont(\PHPExcel_Style_Font $source, \PHPExcel_Style_Font $dest)
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
     * @param \PHPExcel_Style_NumberFormat $source  The source
     * @param \PHPExcel_Style_NumberFormat $dest  The destination
     */
    public static function copyNumberFormat(\PHPExcel_Style_NumberFormat $source, \PHPExcel_Style_NumberFormat $dest)
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
     * @param \PHPExcel_Style $source  The source
     * @param \PHPExcel_Style $dest  The destination
     */
    public static function copyStyle(\PHPExcel_Style $source, \PHPExcel_Style $dest)
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