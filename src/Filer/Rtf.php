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

namespace NTLAB\Report\Filer;

use NTLAB\Script\Core\Script;

/**
 * Quick and dirty filer for RTF using regular expression.
 *
 * @author Toha
 */
class Rtf implements FilerInterface
{
    const TAG_SIGN = '%';

    /**
     * @var \NTLAB\Report\Filer\Rtf
     */
    protected static $instance = null;

    /**
     * @var \NTLAB\Script\Core\Script
     */
    protected $script = null;

    /**
     * @var array
     */
    protected $separatorMaps = array('\\', ' ');

    /**
     * @var array
     */
    protected $cleanMaps = array(
        "\n" => '',
        "\r" => '',
        '}' => '',
        '{' => ''
    );

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
    protected $matches = array();

    /**
     * @var array
     */
    protected $hashes = array();

    /**
     * @var array
     */
    protected $caches = array();

    /**
     * @var array
     */
    protected $cleans = array();

    /**
     * @var string
     */
    protected $content = null;

    /**
     * @var string
     */
    protected $header = null;

    /**
     * @var string
     */
    protected $body = null;

    /**
     * @var string
     */
    protected $footer = null;

    /**
     * Get the instance.
     *
     * @return \NTLAB\Report\Filer\Rtf
     */
    public static function getInstance()
    {
        if (null == static::$instance) {
            static::$instance = new self();
        }

        return static::$instance;
    }

    /**
     * {@inheritDoc}
     * @see \NTLAB\Report\Filer\FilerInterface::getScript()
     */
    public function getScript()
    {
        if (null == $this->script) {
            $this->script = new Script();
        }

        return $this->script;
    }

    /**
     * {@inheritDoc}
     * @see \NTLAB\Report\Filer\FilerInterface::setScript()
     */
    public function setScript(Script $script)
    {
        $this->script = $script;

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
        $stag = substr(static::TAG_SIGN, 0, 1);
        $etag = strlen(static::TAG_SIGN) > 1 ? substr(static::TAG_SIGN, 1, 1) : $stag;

        return array($stag, $etag);
    }

    /**
     * Get tag regular expression pattern.
     *
     * @return string
     */
    public function getTagRegex()
    {
        if (null == $this->tagRe) {
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
        if (null == $this->partRe) {
            $allowedChars = 'A-Za-z0-9\.\_';
            $this->partRe = sprintf('/(?P<VALUE>[%1$s]+)(\%2$s(?P<SUBTYPE>[%1$s\(\)]+))?(\%3$s(?P<OPTIONS>[%1$s\%4$s]+))?(\%5$s(?P<CASE>[%1$s]+))?/x', $allowedChars, self::TAG_SUBTYPE_DELIM, self::TAG_OPTIONS_DELIM, self::TAG_OPTIONS_SPLIT, self::TAG_CASE_DELIM);
        }

        return $this->partRe;
    }

    /**
     * Clean a rtf text and return only plain text.
     *
     * @param string $text  The input text
     * @param bool $cache  Cache content
     * @return string Cleaned text
     */
    public function clean($text, $cache = true)
    {
        if ($text) {
            if ($cache && array_key_exists($text, $this->cleans)) {
                $text = $this->cleans[$text];
            } else {
                if ($cache) {
                    $otext = $text;
                }
                foreach ($this->cleanMaps as $search => $replace) {
                    $text = str_replace($search, $replace, $text);
                }
                while (true) {
                    $s = strpos($text, $this->separatorMaps[0]);
                    $e = strpos($text, $this->separatorMaps[1], $s);
                    if (false !== $s && false !== $e) {
                        $text = substr($text, 0, $s).substr($text, $e + 1);
                    } else {
                        break;
                    }
                }
                if ($cache) {
                    $this->cleans[$otext] = $text;
                }
            }
        }

        return $text;
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
        $tags = array($tag);
        if (preg_match_all($this->getPartRegex(), $tag, $matches)) {
            foreach (array(
                0 => 'VALUE',
                self::TAG_SUBTYPE_DELIM => 'SUBTYPE',
                self::TAG_OPTIONS_DELIM => 'OPTIONS',
                self::TAG_CASE_DELIM => 'CASE'
            ) as $key => $name) {
                if (isset($matches[$name]) && ($value = $matches[$name][0])) {
                    $tags[$key] = $value;
                }
            }
            // fix tags in case only got value
            if (1 == count($tags)) {
                $tags[0] = $tag;
            }
        }
        if ($cache) {
            $this->caches[$tag] = $tags;
        }

        return $tags;
    }

    /**
     * Clear the caches.
     *
     * @return \NTLAB\Report\Filer\Rtf
     */
    public function clear()
    {
        $this->cleans = array();
        $this->caches = array();

        return $this;
    }

    /**
     * Parse a tag.
     *
     * @param string $tag   Tag expression
     * @return string
     */
    protected function parseTag($tag)
    {
        $tag = $this->clean($tag);
        $value = null;
        if (!$this->getScript()->getVar($value, $tag, $this->getScript()->getContext())) {
            $value = $this->getScript()->evaluate($tag);
        }

        return $value;
    }

    /**
     * Get the EACH tag in template.
     *
     * @param string $template  The template
     * @param array $matches  The matches tags
     * @return array
     */
    protected function parseEach(&$template, $matches)
    {
        $eachs = array();
        for ($i = 0; $i < count($matches[1]); $i ++) {
            $match = $matches[0][$i];
            $tag = $this->clean($matches[1][$i]);
            if ('EACH:' == substr($tag, 0, 5)) {
                $tags = explode(':', $tag, 3);
                $eachs[$tags[1]] = array(
                    'start' => $match,
                    'end' => null,
                    'expr' => $tags[2],
                    'content' => null
                );
            }
            if ('EACHE:' == substr($tag, 0, 6)) {
                $tags = explode(':', $tag, 2);
                if (isset($eachs[$tags[1]])) {
                    $eachs[$tags[1]]['end'] = $match;
                }
            }
        }
        $keys = array_keys($eachs);
        for ($i = 0; $i < count($keys); $i ++) {
            if ($eachs[$keys[$i]]['start'] && $eachs[$keys[$i]]['end']) {
                $s = strpos($template, $eachs[$keys[$i]]['start']);
                $e = strpos($template, $eachs[$keys[$i]]['end']);
                if (is_array($parts = $this->splitParts($template, $s, $e, $eachs[$keys[$i]]['start'], $eachs[$keys[$i]]['end']))) {
                    $eachs[$keys[$i]]['content'] = $parts[1];
                    $template = $parts[0].'%%EACH:'.$keys[$i].'%%'.$parts[2];
                }
            }
        }

        return $eachs;
    }

    /**
     * Get the TBL tag in template.
     *
     * @param string $template  The template
     * @param array $matches  The matches tags
     * @return array
     */
    protected function parseTable(&$template, $matches)
    {
        $tables = array();
        for ($i = 0; $i < count($matches[1]); $i ++) {
            $match = $matches[0][$i];
            $tag = $this->clean($matches[1][$i]);
            if ('TBL:' == substr($tag, 0, 4)) {
                $tags = explode(':', $tag, 3);
                $tables[$tags[1]] = array(
                    'start' => $match,
                    'end' => null,
                    'expr' => $tags[2],
                    'content' => null
                );
            }
            if ('TBLE:' == substr($tag, 0, 5)) {
                $tags = explode(':', $tag, 2);
                if (isset($tables[$tags[1]])) {
                    $tables[$tags[1]]['end'] = $match;
                }
            }
        }
        $keys = array_keys($tables);
        for ($i = 0; $i < count($keys); $i ++) {
            if ($tables[$keys[$i]]['start'] && $tables[$keys[$i]]['end']) {
                // find table row begin \trowd
                if (false !== ($s = strpos($template, $tables[$keys[$i]]['start']))) {
                    $s = strrpos(substr($template, 0, $s - 1), '\trowd ');
                }
                // find table row end \row followed by \pard
                if (false !== ($e = strpos($template, $tables[$keys[$i]]['end']))) {
                    if (false !== ($e = strpos($template, '\row ', $e))) {
                        $e = strpos($template, '\pard ', $e);
                    }
                }
                if (is_int($s) && is_int($e)) {
                    $header = substr($template, 0, $s);
                    $content = substr($template, $s, $e - $s);
                    $footer = substr($template, $e);
                    $content = strtr($content, array(
                        $tables[$keys[$i]]['start'] => '',
                        $tables[$keys[$i]]['end'] => ''
                    ));
                    $tables[$keys[$i]]['content'] = $content;
                    $template = $header.'%%TBL:'.$keys[$i].'%%'.$footer;
                }
            }
        }

        return $tables;
    }

    /**
     * Parse and replace all tags in template.
     *
     * @param string $template  The template content
     * @param string $regex  The regular expression pattern
     * @return string
     */
    public function parse($template, $regex = null, $script = false)
    {
        if ($template) {
            // check if template has cached
            $md5 = md5($template);
            if (!in_array($md5, $this->hashes)) {
                $matches = null;
                preg_match_all(null !== $regex ? $regex : $this->getTagRegex(), $template, $matches, PREG_PATTERN_ORDER);
                $this->matches[$md5] = $matches;
                $this->hashes[] = $md5;
            } else {
                $matches = $this->matches[$md5];
            }
            // scan for EACH tags
            $eachs = $this->parseEach($template, $matches);
            // scan for TBL tags
            $tables = $this->parseTable($template, $matches);
            // parse regular tags
            for ($i = 0; $i < count($matches[1]); $i ++) {
                $match = $matches[0][$i];
                $tag = $matches[1][$i];
                // ignore non exist match
                if (false === strpos($template, $match)) {
                    continue;
                }
                if (false !== ($tag = $this->parseTag($tag))) {
                    if ($tag instanceof \DateTime) {
                        $tag = $tag->format(\DateTime::ISO8601);
                    }
                    $template = str_replace($match, $tag, $template);
                }
            }
            // process EACH
            foreach ($eachs as $tag => $params) {
                $content = null;
                if ($params['expr'] && $params['content']) {
                    $this->getScript()->pushContext();
                    $objects = $this->getScript()->evaluate($params['expr']);
                    $filer = new self();
                    $filer->setScript($this->getScript());
                    $content = $filer->build($params['content'], $objects, true);
                    $this->getScript()->popContext();
                }
                $template = str_replace('%%EACH:'.$tag.'%%', $content, $template);
            }
            // process TBL
            foreach ($tables as $tag => $params) {
                $content = null;
                if ($params['expr'] && $params['content']) {
                    $this->getScript()->pushContext();
                    $objects = $this->getScript()->evaluate($params['expr']);
                    $filer = new self();
                    $filer->setScript($this->getScript());
                    $content = $filer->build($params['content'], $objects, true, "\n");
                    $this->getScript()->popContext();
                }
                $template = str_replace('%%TBL:'.$tag.'%%', $content, $template);
            }
        }

        return $template;
    }

    /**
     * Get nearest rtf item.
     *
     * @param string $text  The rtf content
     * @param int $pos  Start position
     * @param int $dir  Direction
     */
    protected function nearestItem($text, &$pos, $dir = -1)
    {
        switch (true) {
            // backward direction
            case $dir < 0:
                $s = substr($text, 0, $pos);
                $bpos = strrpos($s, '{');
                $epos = strrpos($s, '}');
                if (false !== $bpos) {
                    // begin tag is not whithin end tag
                    if (false === $epos || $bpos > $epos) {
                        $pos = $bpos;
                    }
                }
                break;

            // forward direction
            case $dir > 0:
                $bpos = strpos($text, '{');
                $epos = strpos($text, '}');
                if (false !== $epos) {
                    // end pos is not prefixed with begin tag
                    if (false === $bpos || $epos < $bpos) {
                        $pos = $epos + 1;
                    }
                }
                break;
        }
    }

    /**
     * Find the tag in text.
     *
     * @param string $text  The text
     * @param string $regex  The tag to find
     * @param string $match  The text matched the regex
     * @return int|bool The position or false if no match
     */
    protected function findTag($text, $regex, &$match)
    {
        $matches = null;
        $tags = $this->getTags();
        $pattern = sprintf('/%1$s(.*)(%3$s)(.*)%2$s/', $tags[0], $tags[1], $regex);
        if (preg_match($pattern, $text, $matches, PREG_OFFSET_CAPTURE)) {
            $match = $matches[0][0];

            return $matches[0][1];
        }

        return false;
    }

    /**
     * Split text parts.
     *
     * @param string $text  The text
     * @param int $start  The start position
     * @param int $end  The end position
     * @param string $startMatch  The start text matched
     * @param string $endMatch  The end text matched
     * @return array
     */
    protected function splitParts($text, &$start, &$end, $startMatch, $endMatch)
    {
        $s = $start;
        $e = $end;
        if ($e > $s) {
            $this->nearestItem($text, $s);
            $this->nearestItem($text, $e);
            if ($s == $e) {
                $s = $start;
                $e = $end;
            }
            $header = substr($text, 0, $s);
            $footer = str_replace($endMatch, '', substr($text, $e));
            $body = str_replace($startMatch, '', substr($text, $s, $e - $s));

            return array($header, $body, $footer);
        }
    }

    /**
     * Extract a template into parts: header, body, and footer.
     *
     * @param string $template  The template
     * @param string $sregex  The begin tag regex
     * @param string $eregex  The end tag regex
     * @return bool|array If success, return array(header, body, footer), false otherwise
     */
    public function extract($template, $sregex = 'BEGIN', $eregex = 'END')
    {
        if (is_resource($template)) {
            $template = stream_get_contents($template);
        }
        $smatch = null;
        $ematch = null;
        $s = $this->findTag($template, $sregex, $smatch);
        $e = $this->findTag($template, $eregex, $ematch);
        if (is_array($parts = $this->splitParts($template, $s, $e, $smatch, $ematch))) {
            return $parts;
        }

        return false;
    }

    /**
     * Begin document creation.
     *
     * @return \NTLAB\Report\Filer\Rtf
     */
    public function beginDoc()
    {
        $this->content = $this->header;

        return $this;
    }

    /**
     * End document creation.
     *
     * @return \NTLAB\Report\Filer\Rtf
     */
    public function endDoc()
    {
        $this->content .= $this->footer;

        return $this;
    }

    /**
     * Add page break to create a new page.
     *
     * @param bool $new  True to add a page break
     * @param string $type  Break type
     * @return \NTLAB\Report\Filer\Rtf
     */
    public function newPage($new, $break = '\page')
    {
        if ($new) {
            $this->content .= $break;
        }

        return $this;
    }

    /**
     * Create a page and build and replace tags.
     *
     * @return \NTLAB\Report\Filer\Rtf
     */
    public function createPage()
    {
        $this->content .= $this->parse($this->body);

        return $this;
    }

    /**
     * Build data for objects using provided template.
     *
     * @param string $template  The template
     * @param array $objects  The template objects
     * @param bool $child  True to build child objects
     * @param string $separator  Child new page separator
     * @return string
     */
    public function build($template, $objects, $child = false, $separator = null)
    {
        if ($child || is_array($parts = $this->extract($template))) {
            $this->header = $child ? null : $parts[0];
            $this->body = $child ? $template : $parts[1];
            $this->footer = $child ? null : $parts[2];
            $this->beginDoc();
            $this->getScript()
                ->setObjects($objects)
                ->each(function(Script $script, Rtf $_this) use ($child, $separator) {
                    $_this
                        ->newPage($script->getIterator()->getRecNo() > 1, $child ? (null !== $separator ? $separator : '\par') : null)
                        ->createPage();
                }, !$child)
            ;
            $this->endDoc();

            return $this->content;
        }
    }

    /**
     * Get the generated template content.
     *
     * @return string
     */
    public function __toString()
    {
        return $this->content;
    }
}