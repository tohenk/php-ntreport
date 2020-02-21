<?php

/*
 * The MIT License
 *
 * Copyright (c) 2014-2020 Toha <tohenk@yahoo.com>
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

use NTLAB\Report\Util\Tag;
use NTLAB\Script\Core\Script;

/**
 * RTF Tag filer using regular expression.
 *
 * @author Toha
 */
class RtfTag implements FilerInterface
{
    /**
     * @var \NTLAB\Report\Util\Tag
     */
    protected $tag = null;

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
    protected $cleans = array();

    /**
     * @var string
     */
    protected $content = null;

    /**
     * @var string
     */
    protected $template = null;

    /**
     * @var string
     */
    protected $break = '\par';

    /**
     * @var boolean
     */
    protected $notifyContextChange = false;

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
     * Get tag.
     *
     * @return \NTLAB\Report\Util\Tag
     */
    public function getTag()
    {
        if (null === $this->tag) {
            $this->tag = new Tag();
        }

        return $this->tag;
    }

    /**
     * Clear the caches.
     *
     * @return \NTLAB\Report\Filer\RtfTag
     */
    public function clear()
    {
        if ($this->tag) {
            $this->tag->clear();
        }
        $this->cleans = array();

        return $this;
    }

    /**
     * Clean a rtf text and return only plain text.
     *
     * @param string $text  The input text
     * @param bool $cache  Cache content
     * @return string Cleaned text
     */
    protected function clean($text, $cache = true)
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
     * Parse sub template.
     *
     * @param string $template  Template content
     * @param string $expr      Object expression
     * @param string $break     Break separator
     * @return string
     */
    protected function parseSub($template, $expr, $break = null)
    {
        $content = null;
        if (strlen($template) && strlen($expr)) {
            $this->getScript()->pushContext();
            try {
                if (null !== ($objects = $this->getScript()->evaluate($expr))) {
                    $filer = new self();
                    $filer->setScript($this->getScript());
                    if (null !== $break) {
                        $filer->break = $break;
                    }
                    $content = $filer->build($template, $objects);
                    unset($filer);
                } else {
                    error_log(sprintf('Expression "%s" return NULL!', $expr));
                }
            }
            catch (\Exception $e) {
                $message = null;
                while (null !== $e) {
                    if (null === $message) {
                        $message = $e->getMessage();
                    } else {
                        $message = sprintf('%s [%s]', $message, $e->getMessage());
                    }
                    $e = $e->getPrevious();
                }
                error_log($message);
            }
            $this->getScript()->popContext();
        }
  
        return $content;
    }

    /**
     * Parse and replace all tags in template.
     *
     * @param string $template  The template content
     * @return string
     */
    protected function parse($template)
    {
        if ($template) {
            $regex = $this->getTag()->getTagRegex();
            // check if template has cached
            $md5 = md5($template);
            if (!in_array($md5, $this->hashes)) {
                $matches = null;
                preg_match_all($regex, $template, $matches, PREG_PATTERN_ORDER);
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
                $content = $this->parseSub($params['content'], $params['expr']);
                $template = str_replace('%%EACH:'.$tag.'%%', $content, $template);
            }
            // process TBL
            foreach ($tables as $tag => $params) {
                $content = $this->parseSub($params['content'], $params['expr'], "\n");
                $template = str_replace('%%TBL:'.$tag.'%%', $content, $template);
            }
        }

        return $template;
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
        $tags = $this->getTag()->getTags();
        $pattern = sprintf('/%1$s(.*)(%3$s)(.*)%2$s/', $tags[0], $tags[1], $regex);
        if (preg_match($pattern, $text, $matches, PREG_OFFSET_CAPTURE)) {
            $match = $matches[0][0];

            return $matches[0][1];
        }

        return false;
    }

    /**
     * Begin document creation.
     *
     * @return \NTLAB\Report\Filer\RtfTag
     */
    protected function beginDoc()
    {
        $this->content = null;

        return $this;
    }

    /**
     * End document creation.
     *
     * @return \NTLAB\Report\Filer\RtfTag
     */
    protected function endDoc()
    {
        return $this;
    }

    /**
     * Add break.
     *
     * @param bool $new  True to add a break
     * @return \NTLAB\Report\Filer\RtfTag
     */
    protected function addBreak($new)
    {
        if ($new) {
            $this->content .= $this->break;
        }

        return $this;
    }

    /**
     * Create content by replace tags.
     *
     * @return \NTLAB\Report\Filer\RtfTag
     */
    protected function addContent()
    {
        $this->content .= $this->parse($this->template);

        return $this;
    }

    /**
     * {@inheritDoc}
     * @see \NTLAB\Report\Filer\FilerInterface::build()
     */
    public function build($template, $objects)
    {
        if ($template) {
            $this->template = $template;
            $this->beginDoc();
            $this->getScript()
                ->setObjects($objects)
                ->each(function(Script $script, RtfTag $_this) {
                    $_this
                        ->addBreak($script->getIterator()->getRecNo() > 1)
                        ->addContent();
                }, $this->notifyContextChange)
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