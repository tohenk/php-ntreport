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

namespace NTLAB\Report\Script;

use NTLAB\Script\Core\Module;
use NTLAB\Report\Util\Spreadsheet\RichText;

/**
 * Excel report functions.
 *
 * @author Toha
 * @id report.excel
 */
class ReportExcel extends Module
{
    /**
     * Make `text` bold.
     *
     * @param string $text  The text
     * @return string
     * @func xlbold
     */
    public function f_XlsRtBold($text)
    {
        if ($text) {
            $text = RichText::tag(RichText::TAG_BOLD, $text);
        }

        return $text;
    }

    /**
     * Make `text` italic.
     *
     * @param string $text  The text
     * @return string
     * @func xlitalic
     */
    public function f_XlsRtItalic($text)
    {
        if ($text) {
            $text = RichText::tag(RichText::TAG_ITALIC, $text);
        }

        return $text;
    }

    /**
     * Make `text` underlined.
     *
     * @param string $text  The text
     * @return string
     * @func xlunderline
     */
    public function f_XlsRtUnderline($text)
    {
        if ($text) {
            $text = RichText::tag(RichText::TAG_UNDERLINE, $text);
        }

        return $text;
    }

    /**
     * Make `text` strike-through.
     *
     * @param string $text  The text
     * @return string
     * @func xlstrike
     */
    public function f_XlsRtStrikethrough($text)
    {
        if ($text) {
            $text = RichText::tag(RichText::TAG_STRIKETHROUGH, $text);
        }

        return $text;
    }

    /**
     * Make `text` subscript.
     *
     * @param string $text  The text
     * @return string @func xlsub
     */
    public function f_XlsRtSubscript($text)
    {
        if ($text) {
            $text = RichText::tag(RichText::TAG_SUBSCRIPT, $text);
        }

        return $text;
    }

    /**
     * Make `text` superscript.
     *
     * @param string $text  `The text
     * @return string
     * @func xlsup
     */
    public function f_XlsRtSuperscript($text)
    {
        if ($text) {
            $text = RichText::tag(RichText::TAG_SUPERSCRIPT, $text);
        }
        
        return $text;
    }

    /**
     * Make `text` sized at `size`.
     *
     * @param string $text
     *            The text
     * @param int $size
     *            The size
     * @return string @func xlsz
     */
    public function f_XlsRtSize($text, $size)
    {
        if ($text) {
            $text = RichText::tag(RichText::TAG_SIZE, $text, $size);
        }

        return $text;
    }

    /**
     * Set `text` color to `color`.
     *
     * @param string $text  The text
     * @param string $color  The color
     * @return string
     * @func xlcolor
     */
    public function f_XlsRtColor($text, $color)
    {
        if ($text) {
            $text = RichText::tag(RichText::TAG_COLOR, $text, $color);
        }

        return $text;
    }
}