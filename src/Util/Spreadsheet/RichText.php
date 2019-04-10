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

use PhpOffice\PhpSpreadsheet\RichText as XlRichText;
use PhpOffice\PhpSpreadsheet\RichText\Run as XlRichTextRun;

class RichText
{
    const TAG_BOLD = 'b';
    const TAG_ITALIC = 'i';
    const TAG_UNDERLINE = 'u';
    const TAG_STRIKETHROUGH = 'strike';
    const TAG_SUBSCRIPT = 'sub';
    const TAG_SUPERSCRIPT = 'sup';
    const TAG_SIZE = 'size';
    const TAG_COLOR = 'color';

    protected static $tags = array(
        self::TAG_BOLD,
        self::TAG_ITALIC,
        self::TAG_ITALIC,
        self::TAG_STRIKETHROUGH,
        self::TAG_SUBSCRIPT,
        self::TAG_SUPERSCRIPT,
        self::TAG_SIZE,
        self::TAG_COLOR
    );

    protected static $colors = array(
        'aliceblue' => 'FFF0F8FF',
        'antiquewhite' => 'FFFAEBD7',
        'aqua' => 'FF00FFFF',
        'aquamarine' => 'FF7FFFD4',
        'azure' => 'FFF0FFFF',
        'beige' => 'FFF5F5DC',
        'bisque' => 'FFFFE4C4',
        'black' => 'FF000000',
        'blanchedalmond' => 'FFFFEBCD',
        'blue' => 'FF0000FF',
        'blueviolet' => 'FF8A2BE2',
        'brown' => 'FFA52A2A',
        'burlywood' => 'FFDEB887',
        'cadetblue' => 'FF5F9EA0',
        'chartreuse' => 'FF7FFF00',
        'chocolate' => 'FFD2691E',
        'coral' => 'FFFF7F50',
        'cornflowerblue' => 'FF6495ED',
        'cornsilk' => 'FFFFF8DC',
        'crimson' => 'FFDC143C',
        'cyan' => 'FF00FFFF',
        'darkblue' => 'FF00008B',
        'darkcyan' => 'FF008B8B',
        'darkgoldenrod' => 'FFB8860B',
        'darkgray' => 'FFA9A9A9',
        'darkgreen' => 'FF006400',
        'darkgrey' => 'FFA9A9A9',
        'darkkhaki' => 'FFBDB76B',
        'darkmagenta' => 'FF8B008B',
        'darkolivegreen' => 'FF556B2F',
        'darkorange' => 'FFFF8C00',
        'darkorchid' => 'FF9932CC',
        'darkred' => 'FF8B0000',
        'darksalmon' => 'FFE9967A',
        'darkseagreen' => 'FF8FBC8F',
        'darkslateblue' => 'FF483D8B',
        'darkslategray' => 'FF2F4F4F',
        'darkslategrey' => 'FF2F4F4F',
        'darkturquoise' => 'FF00CED1',
        'darkviolet' => 'FF9400D3',
        'deeppink' => 'FFFF1493',
        'deepskyblue' => 'FF00BFFF',
        'dimgray' => 'FF696969',
        'dimgrey' => 'FF696969',
        'dodgerblue' => 'FF1E90FF',
        'firebrick' => 'FFB22222',
        'floralwhite' => 'FFFFFAF0',
        'forestgreen' => 'FF228B22',
        'fuchsia' => 'FFFF00FF',
        'gainsboro' => 'FFDCDCDC',
        'ghostwhite' => 'FFF8F8FF',
        'gold' => 'FFFFD700',
        'goldenrod' => 'FFDAA520',
        'gray' => 'FF808080',
        'green' => 'FF008000',
        'greenyellow' => 'FFADFF2F',
        'grey' => 'FF808080',
        'honeydew' => 'FFF0FFF0',
        'hotpink' => 'FFFF69B4',
        'indianred' => 'FFCD5C5C',
        'indigo' => 'FF4B0082',
        'ivory' => 'FFFFFFF0',
        'khaki' => 'FFF0E68C',
        'lavender' => 'FFE6E6FA',
        'lavenderblush' => 'FFFFF0F5',
        'lawngreen' => 'FF7CFC00',
        'lemonchiffon' => 'FFFFFACD',
        'lightblue' => 'FFADD8E6',
        'lightcoral' => 'FFF08080',
        'lightcyan' => 'FFE0FFFF',
        'lightgoldenrodyellow' => 'FFFAFAD2',
        'lightgray' => 'FFD3D3D3',
        'lightgreen' => 'FF90EE90',
        'lightgrey' => 'FFD3D3D3',
        'lightpink' => 'FFFFB6C1',
        'lightsalmon' => 'FFFFA07A',
        'lightseagreen' => 'FF20B2AA',
        'lightskyblue' => 'FF87CEFA',
        'lightslategray' => 'FF778899',
        'lightslategrey' => 'FF778899',
        'lightsteelblue' => 'FFB0C4DE',
        'lightyellow' => 'FFFFFFE0',
        'ltgray' => 'FFC0C0C0',
        'medgray' => 'FFA0A0A0',
        'dkgray' => 'FF808080',
        'moneygreen' => 'FFC0DCC0',
        'legacyskyblue' => 'FFF0CAA6',
        'cream' => 'FFF0FBFF',
        'lime' => 'FF00FF00',
        'limegreen' => 'FF32CD32',
        'linen' => 'FFFAF0E6',
        'magenta' => 'FFFF00FF',
        'maroon' => 'FF800000',
        'mediumaquamarine' => 'FF66CDAA',
        'mediumblue' => 'FF0000CD',
        'mediumorchid' => 'FFBA55D3',
        'mediumpurple' => 'FF9370DB',
        'mediumseagreen' => 'FF3CB371',
        'mediumslateblue' => 'FF7B68EE',
        'mediumspringgreen' => 'FF00FA9A',
        'mediumturquoise' => 'FF48D1CC',
        'mediumvioletred' => 'FFC71585',
        'midnightblue' => 'FF191970',
        'mintcream' => 'FFF5FFFA',
        'mistyrose' => 'FFFFE4E1',
        'moccasin' => 'FFFFE4B5',
        'navajowhite' => 'FFFFDEAD',
        'navy' => 'FF000080',
        'oldlace' => 'FFFDF5E6',
        'olive' => 'FF808000',
        'olivedrab' => 'FF6B8E23',
        'orange' => 'FFFFA500',
        'orangered' => 'FFFF4500',
        'orchid' => 'FFDA70D6',
        'palegoldenrod' => 'FFEEE8AA',
        'palegreen' => 'FF98FB98',
        'paleturquoise' => 'FFAFEEEE',
        'palevioletred' => 'FFDB7093',
        'papayawhip' => 'FFFFEFD5',
        'peachpuff' => 'FFFFDAB9',
        'peru' => 'FFCD853F',
        'pink' => 'FFFFC0CB',
        'plum' => 'FFDDA0DD',
        'powderblue' => 'FFB0E0E6',
        'purple' => 'FF800080',
        'red' => 'FFFF0000',
        'rosybrown' => 'FFBC8F8F',
        'royalblue' => 'FF4169E1',
        'saddlebrown' => 'FF8B4513',
        'salmon' => 'FFFA8072',
        'sandybrown' => 'FFF4A460',
        'seagreen' => 'FF2E8B57',
        'seashell' => 'FFFFF5EE',
        'sienna' => 'FFA0522D',
        'silver' => 'FFC0C0C0',
        'skyblue' => 'FF87CEEB',
        'slateblue' => 'FF6A5ACD',
        'slategray' => 'FF708090',
        'slategrey' => 'FF708090',
        'snow' => 'FFFFFAFA',
        'springgreen' => 'FF00FF7F',
        'steelblue' => 'FF4682B4',
        'tan' => 'FFD2B48C',
        'teal' => 'FF008080',
        'thistle' => 'FFD8BFD8',
        'tomato' => 'FFFF6347',
        'turquoise' => 'FF40E0D0',
        'violet' => 'FFEE82EE',
        'wheat' => 'FFF5DEB3',
        'white' => 'FFFFFFFF',
        'whitesmoke' => 'FFF5F5F5',
        'yellow' => 'FFFFFF00',
        'yellowgreen' => 'FF9ACD32',
    );

    /**
     * Get tag regular expression.
     *
     * @param string $tag  The tag
     * @return string
     */
    protected static function tagRegex($tag)
    {
        return sprintf('/\<%1$s(\:([a-zA-Z0-9]+))*\>(.*?)\<\/%1$s\>/', $tag);
    }

    /**
     * Get the tag offsets ocurrance.
     *
     * @param string $text  The text to search
     * @return array
     */
    protected static function tagOffset($text)
    {
        $offsets = array();
        foreach (self::$tags as $tag) {
            $regex = self::tagRegex($tag);
            if (preg_match_all($regex, $text, $matches, PREG_OFFSET_CAPTURE)) {
                for ($i = 0; $i < count($matches[0]); $i ++) {
                    $offset = $matches[0][$i][1];
                    $match = $matches[0][$i][0];
                    $extra = $matches[2][$i][0];
                    $inner = $matches[3][$i][0];
                    $offsets[$offset] = array(
                        'tag' => $tag,
                        'match' => $match,
                        'ofs' => $offset,
                        'len' => strlen($match),
                        'extra' => $extra,
                        'inner' => $inner
                    );
                }
            }
        }
        ksort($offsets);

        return $offsets;
    }

    /**
     * Clean tag offsets and return only the outer one.
     *
     * @param array $offsets  The offsets
     * @return array
     */
    protected static function cleanOffset($offsets = array())
    {
        if (count($offsets) > 1) {
            $result = array();
            $result[] = array_shift($offsets);
            foreach ($offsets as $data) {
                foreach ($result as $ref) {
                    if ($data['ofs'] < $ref['ofs'] + $ref['len']) {
                        continue 2;
                    }
                }
                $result[] = $data;
            }

            return $result;
        } else {
            return $offsets;
        }
    }

    /**
     * Clean tags from text.
     *
     * @param string $text  The text
     * @param array $options  Applied tags
     */
    protected static function cleanText(&$text, &$options = array())
    {
        if ($text) {
            foreach (self::$tags as $tag) {
                $regex = self::tagRegex($tag);
                if (preg_match_all($regex, $text, $matches, PREG_OFFSET_CAPTURE)) {
                    if ($text == $matches[0][0][0]) {
                        $text = $matches[3][0][0];
                        if (!array_key_exists($tag, $options)) {
                            $options[$tag] = $matches[2][0][0] ? $matches[2][0][0] : null;
                        }
                        self::cleanText($text, $options);
                        break;
                    }
                }
            }
        }
    }

    /**
     * Create text run.
     *
     *
     * @param XlRichText $richtext  The rich text object
     * @param string $text  The text
     * @param array $options  Rich text font options
     * @return xlRichTextRun
     */
    protected static function createTextRun(XlRichText $richtext, $text, $options = array())
    {
        self::cleanText($text, $options);
        $trun = $richtext->createTextRun($text);
        foreach ($options as $tag => $extra) {
            switch ($tag) {
                case self::TAG_BOLD:
                    $trun->getFont()->setBold(true);
                    break;

                case self::TAG_ITALIC:
                    $trun->getFont()->setItalic(true);
                    break;

                case self::TAG_UNDERLINE:
                    $trun->getFont()->setUnderline(true);
                    break;

                case self::TAG_STRIKETHROUGH:
                    $trun->getFont()->setStrikethrough(true);
                    break;

                case self::TAG_SUBSCRIPT:
                    $trun->getFont()->setSubScript(true);
                    break;

                case self::TAG_SUPERSCRIPT:
                    $trun->getFont()->setSuperScript(true);
                    break;

                case self::TAG_SIZE:
                    if (is_numeric($extra)) {
                        $trun->getFont()->setSize(floatval($extra));
                    }
                    break;

                case self::TAG_COLOR:
                    if (null !== $extra) {
                        $color = array_key_exists($extra, self::$colors) ? self::$colors[$extra] : $extra;
                        try {
                            $trun->getFont()
                                ->getColor()
                                ->setARGB($color);
                        } catch (\Exception $e) {
                        }
                    }
                    break;
            }
        }

        return $trun;
    }

    /**
     * Create rich text by replacing tags: <b>, <i>, <u> and others.
     *
     * @param string $text  The plain text
     * @return XlRichText
     */
    public static function create($text)
    {
        if ($text && count($offsets = self::tagOffset($text))) {
            $richText = new XlRichText();
            $outers = self::cleanOffset($offsets);
            $pos = 0;
            foreach ($outers as $offset) {
                // cut plain text
                if ($plain = substr($text, $pos, $offset['ofs'] - $pos)) {
                    $richText->createText($plain);
                }
                // next pos
                $pos = $offset['ofs'] + $offset['len'];
                // rich text font options
                $options = array(
                    $offset['tag'] => $offset['extra'] ? $offset['extra'] : null
                );
                // create text run
                self::createTextRun($richText, $offset['inner'], $options);
            }
            // remaining plain text
            if ($pos < strlen($text)) {
                $richText->createText(substr($text, $pos));
            }

            return $richText;
        }

        return $text;
    }

    /**
     * Create tag.
     *
     * @param string $tag  The tag
     * @param string $text  The text
     * @param string $extra  The extra tag
     * @return string
     */
    public static function tag($tag, $text, $extra = null)
    {
        return sprintf('<%1$s%3$s>%2$s</%1$s>', $tag, $text, $extra ? ':' . $extra : '');
    }
}