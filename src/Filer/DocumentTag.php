<?php

/*
 * The MIT License
 *
 * Copyright (c) 2014-2021 Toha <tohenk@yahoo.com>
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
use PhpOffice\PhpWord\TemplateProcessor;

/**
 * Document Filer using PHPWord Template Processor.
 *
 * @author Toha
 */
class DocumentTag extends TemplateProcessor implements FilerInterface
{
    const DOC_DOCUMENT    = 1;
    const DOC_TABLE       = 2;
    const DOC_EACH        = 3;

    const DOC_SUB_TABLE   = 'TBL';
    const DOC_SUB_EACH    = 'EACH';
    const DOC_SUB_IF      = 'IF';

    /**
     * @var \NTLAB\Report\Util\Tag
     */
    protected $tag = null;

    /**
     * @var \NTLAB\Script\Core\Script
     */
    protected $script = null;

    /**
     * @var int
     */
    protected $docType = self::DOC_DOCUMENT;

    /**
     * @var string
     */
    protected $docTemplate = null;

    /**
     * @var array
     */
    protected $contents = [];

    /**
     * @var array
     */
    protected $vars = [];

    /**
     * @var array
     */
    protected $tags = [];

    /**
     * @var array
     */
    protected $ifs = [];

    /**
     * @var array
     */
    protected $eaches = [];

    /**
     * @var array
     */
    protected $tables = [];

    /**
     * @var array
     */
    protected $cleanMaps = ["#<(/)*(.*?)(/)*>#"];

    /**
     * @var array
     */
    protected $cleans = [];

    /**
     * Constructor.
     *
     * @param string $documentTemplate
     */
    public function __construct($documentTemplate = null)
    {
        if (null !== $documentTemplate) {
            parent::__construct($documentTemplate);
        }
    }

    /**
     * {@inheritDoc}
     * @see \NTLAB\Report\Filer\FilerInterface::getScript()
     */
    public function getScript()
    {
        if (null === $this->script) {
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
     * @return \NTLAB\Report\Filer\DocumentTag
     */
    public function clear()
    {
        if ($this->tag) {
            $this->tag->clear();
        }
        $this->cleans = [];
        return $this;
    }

    /**
     * Clean document element and return only plain text.
     *
     * @param string $text  The input text
     * @param bool $cache  Cache content
     * @return string Cleaned text
     */
    protected function clean($text, $cache = true)
    {
        if (null !== $text) {
            if ($cache && array_key_exists($text, $this->cleans)) {
                $text = $this->cleans[$text];
            } else {
                if ($cache) {
                    $otext = $text;
                }
                foreach ($this->cleanMaps as $pattern) {
                    $matches = [];
                    preg_match_all($pattern, $text, $matches, PREG_OFFSET_CAPTURE);
                    if (count($matches[0])) {
                        $endPos = null;
                        $startPos = null;
                        $cleaned = false;
                        $i = count($matches[0]);
                        while (true) {
                            $i--;
                            if ($i >= 0) {
                                $pos = $matches[0][$i];
                                $s = $pos[1];
                                $e = $s + strlen($pos[0]);
                                if (null === $endPos || null === $startPos) {
                                    $startPos = $s;
                                    $endPos = $e;
                                } else {
                                    // is cleaning necessary
                                    if ($e === $startPos) {
                                        $startPos = $s;
                                        if ($i == 0) {
                                            $cleaned = true;
                                        }
                                    } else {
                                        $cleaned = true;
                                    }
                                }
                            }
                            if ($cleaned && $startPos !== null && $endPos !== null) {
                                $text = $this->replaceStrAt($text, $startPos, $endPos);
                                $startPos = null;
                                $endPos = null;
                                $cleaned = false;
                                // restart from previous
                                if ($i > 0) {
                                    $i++;
                                }
                            }
                            if ($i == 0) {
                                break;
                            }
                        }
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
     * Clean part of string from start to end position.
     *
     * @param string $str
     * @param int $start
     * @param int $end
     * @param string $replacement
     * @return string
     */
    protected function replaceStrAt($str, $start, $end, $replacement = null)
    {
        $sEnd = substr($str, $end);
        $sStart = substr($str, 0, $start);
        return $sStart.$replacement.$sEnd;
    }

    /**
     * Parse a tag.
     *
     * @param string $tag   Tag expression
     * @return string
     */
    protected function parseTag($tag)
    {
        $value = null;
        $tag = $this->clean($tag);
        if (!$this->getScript()->getVar($value, $tag, $this->getScript()->getContext())) {
            $value = $this->getScript()->evaluate($tag);
        }
        return $value;
    }

    /**
     * Create sub tag placeholder.
     *
     * @param string $tag
     * @param string $key
     * @return string
     */
    protected function createSubTag($tag, $key)
    {
        return '%%'.sprintf('%s:%s', $tag, $key).'%%';
    }

    /**
     * Get sub tags in template.
     *
     * @param string $key  Tag key
     * @param array $vars  The variables tags
     * @return array
     */
    protected function findSubTags($key, $vars)
    {
        $items = [];
        $skey = sprintf('%s:', $key);
        $ekey = sprintf('%sE:', $key);
        for ($i = 0; $i < count($vars); $i ++) {
            $tag = $this->clean($vars[$i]);
            $match = $this->getTag()->createTag($vars[$i]);
            if ($skey == substr($tag, 0, strlen($skey))) {
                $tags = explode(':', $tag, 3);
                $items[$tags[1]] = [
                    'start' => $match,
                    'end' => null,
                    'expr' => $tags[2],
                    'content' => null
                ];
            }
            if ($ekey == substr($tag, 0, strlen($ekey))) {
                $tags = explode(':', $tag, 2);
                if (isset($items[$tags[1]])) {
                    $items[$tags[1]]['end'] = $match;
                }
            }
        }
        $keys = array_keys($items);
        for ($i = 0; $i < count($keys); $i ++) {
            if ($items[$keys[$i]]['start'] && $items[$keys[$i]]['end']) {
                if (
                    ($startPos = $this->findContainingXmlBlockForMacro($items[$keys[$i]]['start'])) &&
                    ($endPos = $this->findContainingXmlBlockForMacro($items[$keys[$i]]['end']))
                ) {
                    $items[$keys[$i]]['matchStart'] = substr($this->tempDocumentMainPart,
                        $startPos['start'],
                        $startPos['end'] - $startPos['start']);
                    $items[$keys[$i]]['matchEnd'] = substr($this->tempDocumentMainPart,
                        $endPos['start'],
                        $endPos['end'] - $endPos['start']);
                    $items[$keys[$i]]['content'] = substr($this->tempDocumentMainPart,
                        $startPos['end'],
                        $endPos['start'] - $startPos['end']);
                    $this->tempDocumentMainPart = substr($this->tempDocumentMainPart, 0, $startPos['start']).
                        $this->createSubTag($key, $keys[$i]).
                        substr($this->tempDocumentMainPart, $endPos['end']);
                }
            }
        }
        return $items;
    }

    /**
     * Parse sub template.
     *
     * @param string $template  Template content
     * @param string $expr  Script expression
     * @param int $docType  Document type to parse
     * @return string
     */
    protected function parseSub($template, $expr, $docType)
    {
        $content = null;
        $this->getScript()->pushContext();
        try {
            $filer = new self();
            $filer->setScript($this->getScript());
            $filer->tempDocumentFilename = tempnam(dirname($this->tempDocumentFilename), 'sub');
            $filer->tempDocumentMainPart = $template;
            $filer->docType = $docType;
            if (null !== ($objects = $this->getScript()->evaluate($expr))) {
                $content = $filer->build(null, $objects);
            } else {
                error_log(sprintf('Expression "%s" from %s return NULL!', $expr, $this->getScript()->getContext()));
            }
            @unlink($filer->tempDocumentFilename);
            unset($filer);
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
        return $content;
    }

    /**
     * Restore tag into original.
     *
     * @param string $template
     * @param string $type
     * @param array $tags
     */
    protected function restoreTemplateTag(&$template, $type, $tags)
    {
        foreach ($tags as $tag => $params) {
            $placeholder = $this->createSubTag($type, $tag);
            if (false !== strpos($template, $placeholder)) {
                $content = implode('', [$params['matchStart'], $params['content'], $params['matchEnd']]);
                $template = str_replace($placeholder, $content, $template);
            }
        }
    }

    /**
     * Restore template with special tag original content.
     *
     * @param string $template
     * @return string
     */
    protected function restoreTemplate($template)
    {
        // replace IF
        $this->restoreTemplateTag($template, static::DOC_SUB_IF, $this->ifs);
        // replace TBL
        $this->restoreTemplateTag($template, static::DOC_SUB_TABLE, $this->tables);
        // replace EACH
        $this->restoreTemplateTag($template, static::DOC_SUB_EACH, $this->eaches);
        return $template;
    }

    /**
     * Parse and replace all tags in template.
     *
     * @return \NTLAB\Report\Filer\DocumentTag
     */
    protected function parse()
    {
        if (count($this->vars)) {
            $template = $this->tempDocumentMainPart;
            // process IF
            foreach ($this->ifs as $tag => $params) {
                $placeholder = $this->createSubTag(static::DOC_SUB_IF, $tag);
                if (false === strpos($template, $placeholder)) {
                    continue;
                }
                $content = null;
                if ($params['expr'] && $params['content'] && $this->getScript()->evaluate($params['expr'])) {
                    $content = $params['content'];
                }
                $template = str_replace($placeholder, $content, $template);
            }
            // process TBL
            foreach ($this->tables as $tag => $params) {
                $placeholder = $this->createSubTag(static::DOC_SUB_TABLE, $tag);
                if (false === strpos($template, $placeholder)) {
                    continue;
                }
                $content = null;
                if ($params['expr'] && $params['content']) {
                    $content = $this->parseSub($this->restoreTemplate($params['content']), $params['expr'], static::DOC_TABLE);
                }
                $template = str_replace($placeholder, $content, $template);
            }
            // process EACH
            foreach ($this->eaches as $tag => $params) {
                $placeholder = $this->createSubTag(static::DOC_SUB_EACH, $tag);
                if (false === strpos($template, $placeholder)) {
                    continue;
                }
                $content = null;
                if ($params['expr'] && $params['content']) {
                    $content = $this->parseSub($params['content'], $params['expr'], static::DOC_EACH);
                }
                $template = str_replace($placeholder, $content, $template);
            }
            // process regular tags
            for ($i = 0; $i < count($this->vars); $i ++) {
                $tag = $this->vars[$i];
                $match = $this->getTag()->createTag($tag);
                // ignore non exist match
                if (false === strpos($template, $match)) {
                    continue;
                }
                if (false !== ($tag = $this->parseTag($tag))) {
                    if ($tag instanceof \DateTime) {
                        $tag = $tag->format(\DateTime::ISO8601);
                    }
                    $tag = $this->ensureUtf8Encoded(htmlspecialchars($tag));
                    $this->replaceText($template, $match, $tag);
                }
            }
            $this->contents[] = $template;
        }
    }

    /**
     * Replace text and preserve space if needed.
     *
     * @param string $template
     * @param string $search
     * @param string $replace
     * @return \NTLAB\Report\Filer\DocumentTag
     */
    protected function replaceText(&$template, $search, $replace)
    {
        $len = strlen($search);
        $sText = '<w:t>';
        $eText = '</w:t>';
        $sLen = strlen($sText);
        $eLen = strlen($eText);
        while (true) {
            if (false === ($pos = strpos($template, $search))) {
                break;
            }
            $template = $this->replaceStrAt($template, $pos, $pos + $len, $replace);
            if (
                false !== ($start = strrpos(substr($template, 0, $pos), $sText)) &&
                false !== ($end = strpos($template, $eText, $start))
            ) {
                $content = substr($template, $start + $sLen, $end - $start - $sLen);
                $encoded = $this->getEncodedText($content);
                if ($content !== $encoded) {
                    $template = $this->replaceStrAt($template, $start, $end + $eLen, $encoded);
                }
            }
        }
        return $this;
    }

    /**
     * Get encoded text.
     *
     * @param string $text
     * @return string
     */
    protected function getEncodedText($text)
    {
        $result = [];
        $parts = [];
        while (true) {
            if (!strlen($text)) {
                break;
            }
            $str = $text;
            $tab = false;
            if (false !== ($p = strpos($text, "\t"))) {
                $str = substr($text, 0, $p);
                $tab = true;
            }
            if ($len = strlen($str)) {
                $parts[] = $str;
                $text = substr($text, $len);
            }
            if ($tab) {
                $parts[] = "\t";
                $text = substr($text, 1);
            }
        }
        foreach ($parts as $part) {
            if ("\t" === $part) {
                $result[] = '<w:tab/>';
            } else if (in_array(' ', [substr($part, 0, 1), substr($part, -1)])) {
                $result[] = sprintf('<w:t xml:space="preserve">%s</w:t>', $part);
            } else if (count($parts) > 1) {
                $result[] = sprintf('<w:t>%s</w:t>', $part);
            } else {
                $result[] = $part;
            }
        }
        return implode('', $result);
    }

    /**
     * Find template rows in table, row must be has a variable to be considered as
     * template row.
     *
     * @return \NTLAB\Report\Filer\DocumentTag
     */
    protected function findTableTemplateRow()
    {
        $startRow = null;
        $endRow = null;
        $matches = [];
        $template = $this->tempDocumentMainPart;
        preg_match_all('#<w:tr(.*?)>(.*?)</w:tr>#', $template, $matches, PREG_OFFSET_CAPTURE);
        foreach ($matches[0] as $row) {
            $this->tempDocumentMainPart = $row[0];
            if (count($this->getVariables())) {
                if (null === $startRow) {
                    $startRow = $row;
                    $endRow = $row;
                } else {
                    $endRow = $row;
                }
            }
        }
        if (null === $startRow || null === $endRow) {
            throw new \InvalidArgumentException('Table has no template row');
        }
        $this->tempDocumentMainPart = substr($template, $startRow[1], $endRow[1] + strlen($endRow[0]) - $startRow[1]);
        $this->docTemplate = str_replace($this->tempDocumentMainPart, '%ROWS%', $template);
        return $this;
    }

    /**
     * Prepare build.
     *
     * @return \NTLAB\Report\Filer\DocumentTag
     */
    protected function prepareBuild()
    {
        switch ($this->docType) {
            case static::DOC_DOCUMENT:
                break;
            case static::DOC_TABLE:
                $this->findTableTemplateRow();
                break;
            case static::DOC_EACH:
                break;
        }
        $this->contents = [];
        $this->vars = $this->getVariables();
        $this->ifs = $this->findSubTags(static::DOC_SUB_IF, $this->vars);
        $this->eaches = $this->findSubTags(static::DOC_SUB_EACH, $this->vars);
        $this->tables = $this->findSubTags(static::DOC_SUB_TABLE, $this->vars);
        return $this;
    }

    /**
     * Finish build.
     *
     * @return \NTLAB\Report\Filer\DocumentTag
     */
    protected function finishBuild()
    {
        $this->tempDocumentMainPart = implode('', $this->contents);
        switch ($this->docType) {
            case static::DOC_DOCUMENT:
                break;
            case static::DOC_TABLE:
                $this->tempDocumentMainPart = str_replace('%ROWS%', $this->tempDocumentMainPart, $this->docTemplate);
                break;
            case static::DOC_EACH:
                break;
        }
        return $this;
    }

    /**
     * {@inheritDoc}
     * @see \NTLAB\Report\Filer\FilerInterface::build()
     */
    public function build($template, $objects)
    {
        if ($this->tempDocumentMainPart) {
            $this->prepareBuild();
            $this->getScript()
                ->setObjects($objects)
                ->each(function(Script $script, DocumentTag $_this) {
                    $_this->parse();
                })
            ;
            $this->finishBuild();
            return $this->tempDocumentMainPart;
        }
    }

    /**
     * Override macro identifier.
     *
     * {@inheritDoc}
     * @see \PhpOffice\PhpWord\TemplateProcessor::getVariablesForPart()
     */
    protected function getVariablesForPart($documentPartXML)
    {
        $matches = [];
        preg_match_all($this->getTag()->getTagRegex(), $documentPartXML, $matches);
        return $matches[1];
    }

    /**
     * Override as internal macro format is unused.
     *
     * @param string $macro
     * @return string
     */
    protected static function ensureMacroCompleted($macro)
    {
        return $macro;
    }
}