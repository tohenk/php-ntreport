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

namespace NTLAB\Report\Util;

/**
 * Report TAG utility.
 *
 * @author Toha <tohenk@yahoo.com>
 */
class Tag
{
    public const TAG_SIGN = '%';

    public const TAG_SUBTYPE_DELIM = '!';
    public const TAG_OPTIONS_DELIM = '?';
    public const TAG_CASE_DELIM = ':';

    public const TAG_OPTIONS_SPLIT = ',';

    /**
     * @var string
     */
    protected $tag = null;

    /**
     * @var string
     */
    protected $tagRe = null;

    /**
     * @var string
     */
    protected $partRe = null;

    /**
     * @var array
     */
    protected $caches = [];

    /**
     * Constructor.
     *
     * @param string $tag
     */
    public function __construct($tag = self::TAG_SIGN)
    {
        $this->tag = $tag;
    }

    /**
     * Clear the caches.
     *
     * @return \NTLAB\Report\Util\Tag
     */
    public function clear()
    {
        $this->caches = [];

        return $this;
    }

    /**
     * Get the enclosing tag signs.
     * The result is an array which contains
     * the opening and closing tag sign array(opening, closing).
     *
     * @return array
     */
    public function getTags()
    {
        $stag = substr($this->tag, 0, 1);
        $etag = strlen($this->tag) > 1 ? substr($this->tag, 1, 1) : $stag;

        return [$stag, $etag];
    }

    /**
     * Get tag regular expression pattern.
     *
     * @return string
     */
    public function getTagRegex()
    {
        if (null === $this->tagRe) {
            $tags = $this->getTags();
            $this->tagRe = sprintf('/%1$s([^%1$s]+)%2$s/', $tags[0], $tags[1]);
        }

        return $this->tagRe;
    }

    /**
     * Get part regular expression pattern.
     *
     * @return string
     */
    public function getPartRegex()
    {
        if (null === $this->partRe) {
            $allowedChars = 'A-Za-z0-9\.\_';
            $this->partRe = sprintf(
                '/(?P<VALUE>[%1$s]+)(\%2$s(?P<SUBTYPE>[%1$s\(\)]+))?(\%3$s(?P<OPTIONS>[%1$s\%4$s]+))?(\%5$s(?P<CASE>[%1$s]+))?/x',
                $allowedChars,
                self::TAG_SUBTYPE_DELIM,
                self::TAG_OPTIONS_DELIM,
                self::TAG_OPTIONS_SPLIT,
                self::TAG_CASE_DELIM
            );
        }

        return $this->partRe;
    }

    /**
     * Create a tag.
     *
     * @param string $tag  The tag
     * @return string
     */
    public function createTag($tag)
    {
        $tags = $this->getTags();

        return $tags[0].$tag.$tags[1];
    }

    /**
     * Split tag into array.
     *
     * @param string $tag  The tag to split
     * @param string $cache  Cache the result
     * @return array (0 => 'The value', '!' => 'subtype', '?' => 'options', ':' => 'case')
     */
    public function splitTag($tag, $cache = true)
    {
        // return cache
        if ($cache && array_key_exists($tag, $this->caches)) {
            return $this->caches[$tag];
        }
        $matches = null;
        $tags = [$tag];
        if (preg_match_all($this->getPartRegex(), $tag, $matches)) {
            foreach ([
                0 => 'VALUE',
                static::TAG_SUBTYPE_DELIM => 'SUBTYPE',
                static::TAG_OPTIONS_DELIM => 'OPTIONS',
                static::TAG_CASE_DELIM => 'CASE'
            ] as $key => $name) {
                if (isset($matches[$name]) && ($value = $matches[$name][0])) {
                    $tags[$key] = $value;
                }
            }
            // fix tags in case only got value
            if (1 === count($tags)) {
                $tags[0] = $tag;
            }
        }
        if ($cache) {
            $this->caches[$tag] = $tags;
        }

        return $tags;
    }
}
