<?php

/*
 * The MIT License
 *
 * Copyright (c) 2014-2024 Toha <tohenk@yahoo.com>
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

use DOMDocument;
use DOMXPath;
use NTLAB\Report\Script\ReportCore;
use NTLAB\Report\Symbol;
use NTLAB\Report\Util\Tag;
use NTLAB\Script\Core\Script;
use PhpOffice\PhpWord\TemplateProcessor;

/**
 * Document Filer using PHPWord Template Processor.
 *
 * @author Toha <tohenk@yahoo.com>
 */
class DocumentTag extends TemplateProcessor implements FilerInterface
{
    public const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
    public const RELATIONSHIPS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
    public const IMAGE_RELATIONSHIP_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

    public const DOC_DOCUMENT = 1;
    public const DOC_TABLE = 2;
    public const DOC_EACH = 3;

    public const DOC_SUB_TABLE = 'TBL';
    public const DOC_SUB_EACH = 'EACH';
    public const DOC_SUB_IF = 'IF';

    /**
     * @var \NTLAB\Report\Filer\DocumentTag
     */
    protected $context = null;

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
    protected $images = [];

    /**
     * @var array
     */
    protected $cleanMaps = ["#<(/)*(.*?)(/)*>#"];

    /**
     * @var array
     */
    protected $cleans = [];

    /**
     * @var string
     */
    protected $extra = null;

    /**
     * @var \DOMDocument
     */
    protected $contentTypes = null;

    /**
     * @var \DOMDocument
     */
    protected $relations = null;

    /**
     * @var array<string>
     */
    protected $relationIds = [];

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
     * Get context.
     *
     * @return \NTLAB\Report\Filer\DocumentTag
     */
    public function getContext()
    {
        return $this->context ? $this->context : $this;
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
                                        if ($i === 0) {
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
                            if ($i === 0) {
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
     * Get the right most position of tags.
     *
     * @param string $str
     * @param array<string> $tags
     * @param bool $exclude
     * @return int|null
     */
    protected function getLastTagPos($str, $tags, $exclude = true)
    {
        $pos = null;
        $tags = is_array($tags) ? $tags : [$tags];
        foreach ($tags as $tag) {
            if (false !== ($p = strrpos($str, $tag))) {
                if (null === $pos || $p > $pos) {
                    $pos = $p + ($exclude ? 0 : strlen($tag));
                }
            }
        }
        return $pos;
    }

    /**
     * Get the left most position of tags.
     *
     * @param string $str
     * @param array<string> $tags
     * @param bool $exclude
     * @return int|null
     */
    protected function getFirstTagPos($str, $tags, $exclude = true)
    {
        $pos = null;
        $tags = is_array($tags) ? $tags : [$tags];
        foreach ($tags as $tag) {
            if (false !== ($p = strpos($str, $tag))) {
                if (null === $pos || $p < $pos) {
                    $pos = $p + ($exclude ? strlen($tag) : 0);
                }
            }
        }
        return $pos;
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
        switch (true) {
            // check if replacement should be replaced with symbol
            case null === $replacement:
                $replaced = substr($str, $start, $end - $start + 1);
                $matches = [];
                preg_match_all('#<w:sym\s(.*?)/>#', $replaced, $matches, PREG_OFFSET_CAPTURE);
                if (count($matches[0])) {
                    $symbols = [];
                    foreach ($matches[0] as $symbol) {
                        $pos = $symbol[1];
                        $symStart = substr($replaced, 0, $pos);
                        $symEnd = substr($replaced, $pos + strlen($symbol[0]));
                        $start = $this->getLastTagPos($symStart, ['<w:r>', '<w:r ']);
                        $end = $this->getFirstTagPos($symEnd, ['</w:r>']);
                        if (null !== $start && null !== $end) {
                            $end += $pos + strlen($symbol[0]);
                            $symbols[] = substr($replaced, $start, $end - $start);
                        }
                    }
                    if (count($symbols)) {
                        $replacement = ReportCore::getReport()->addSymbol(implode('', $symbols));
                    }
                }
                break;
            // check if replacement is empty or symbol
            case '' === $replacement:
            case $replacement instanceof Symbol && '</w:r>' === substr($replacement, -6):
                if (null !== ($pos = $this->getLastTagPos($sStart, ['<w:r>', '<w:r ']))) {
                    $sStart = substr($sStart, 0, $pos);
                }
                if (null !== ($pos = $this->getFirstTagPos($sEnd, ['</w:r>']))) {
                    $sEnd = substr($sEnd, $pos);
                }
                break;
        }
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
     * @param array $vars  The variable tags
     * @return array
     */
    protected function findSubTags($key, $vars)
    {
        $items = [];
        $skey = sprintf('%s:', $key);
        $ekey = sprintf('%sE:', $key);
        for ($i = 0; $i < count($vars); $i++) {
            $tag = $this->clean($vars[$i]);
            $match = $this->getTag()->createTag($vars[$i]);
            if ($skey === substr($tag, 0, strlen($skey))) {
                $tags = explode(':', $tag, 4);
                $items[$tags[1]] = [
                    'start' => $match,
                    'end' => null,
                    'expr' => $tags[2],
                    'content' => null
                ];
                if (count($tags) > 3) {
                    $items[$tags[1]]['extra'] = $tags[3];
                }
            }
            if ($ekey === substr($tag, 0, strlen($ekey))) {
                $tags = explode(':', $tag, 2);
                if (isset($items[$tags[1]])) {
                    $items[$tags[1]]['end'] = $match;
                }
            }
        }
        $keys = array_keys($items);
        for ($i = 0; $i < count($keys); $i++) {
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
     * Find images in template part.
     *
     * @param array The variable tags
     * @return array
     */
    protected function findImages(&$vars)
    {
        $items = [];
        $founds = [];
        for ($i = 0; $i < count($vars); $i++) {
            $tag = $this->getTag()->createTag($vars[$i]);
            if (false !== ($p = strpos($this->tempDocumentMainPart, $tag))) {
                $tmplStart = substr($this->tempDocumentMainPart, 0, $p);
                $tmplEnd = substr($this->tempDocumentMainPart, $p + strlen($tag));
                $s = $this->getFirstTagPos($tmplStart, '<w:drawing>', false);
                $e = $this->getLastTagPos($tmplEnd, '</w:drawing>', false);
                if (null !== $s && null !== $e) {
                    $part = substr($tmplStart, $s).$tag.substr($tmplEnd, 0, $e);
                    $items[$vars[$i]] = [
                        'start' => $s,
                        'content' => $part,
                    ];
                    $founds[] = $i;
                }
            }
        }
        foreach (array_reverse($founds) as $index) {
            unset($vars[$index]);
        }
        return $items;
    }

    /**
     * Parse sub template.
     *
     * @param string $template  Template content
     * @param string $expr  Script expression
     * @param int $docType  Document type to parse
     * @param string $extra  Extra data for parsing
     * @return string
     */
    protected function parseSub($template, $expr, $docType, $extra = null)
    {
        $content = null;
        $this->getScript()->pushContext();
        try {
            $filer = new self();
            $filer->setScript($this->getScript());
            $filer->context = $this->getContext();
            $filer->tempDocumentMainPart = $template;
            $filer->docType = $docType;
            $filer->extra = $extra;
            if (null !== ($objects = $this->getScript()->evaluate($expr))) {
                $content = $filer->build(null, $objects);
            } else {
                error_log(sprintf('Expression "%s" from %s return NULL!', $expr, $this->getScript()->getContext()));
            }
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
                $content = implode('', [$params['matchStart'], $this->restoreTemplate($params['content']), $params['matchEnd']]);
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
        $this->contents[] = $this->parseTemplate($this->tempDocumentMainPart);
        return $this;
    }

    /**
     * Parse and replace all tags in template.
     *
     * @param string $template
     * @return string
     */
    protected function parseTemplate($template)
    {
        // process TBL
        foreach ($this->tables as $tag => $params) {
            $placeholder = $this->createSubTag(static::DOC_SUB_TABLE, $tag);
            if (false === strpos($template, $placeholder)) {
                continue;
            }
            $content = null;
            if ($params['expr'] && $params['content']) {
                $content = $this->parseSub($this->restoreTemplate($params['content']), $params['expr'], static::DOC_TABLE, isset($params['extra']) ? $params['extra'] : null);
            }
            $template = str_replace($placeholder, (string) $content, $template);
        }
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
            $template = str_replace($placeholder, (string) $content, $template);
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
            // empty each result in table column should be replaced with paragraph
            if ('' === $content) {
                $pos = strpos($template, $placeholder);
                $tmplStart = substr($template, 0, $pos);
                $tmplEnd = substr($template, $pos + strlen($placeholder));
                if (substr($tmplStart, -9) === '</w:tcPr>' && substr($tmplEnd, 0, 7) === '</w:tc>') {
                    $content = '<w:p/>';
                }
            }
            $template = str_replace($placeholder, (string) $content, $template);
        }
        // process IMAGES
        foreach ($this->images as $tag => $params) {
            $content = null;
            $tmplImg = $params['content'];
            if (($data = $this->parseTag($tag)) && false !== ($imgdata = @getimagesizefromstring($data))) {
                // update content types
                $extension = $this->getImageExtension($imgdata[2]);
                // update relations
                preg_match('#r:embed="(rId\d+)"#', $tmplImg, $matches);
                list($rid, $imageName) = $this->getImageRelationId($extension, $matches[1]);
                // update media
                /** @var \ZipArchive $zip */
                $zip = $this->getContext()->zipClass;
                $zip->addFromString(sprintf('word/%s', $imageName), $data);
                // update image template
                $content = strtr($tmplImg, [$matches[1] => $rid]);
            }
            $template = str_replace($tmplImg, (string) $content, $template);
        }
        // process regular tags
        for ($i = 0; $i < count($this->vars); $i++) {
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
                if (!$tag instanceof Symbol) {
                    $tag = $this->ensureUtf8Encoded(htmlspecialchars((string) $tag, ENT_NOQUOTES|ENT_SUBSTITUTE));
                }
                $this->replaceText($template, $match, $tag);
            }
        }
        return $template;
    }

    /**
     * Get image extension and update content types if needed.
     *
     * @param int $imageType Image type
     * @return string
     */
    protected function getImageExtension($imageType)
    {
        $extensions = [
            'image/jpeg' => 'jpeg',
            'image/png'  => 'png',
            'image/bmp'  => 'bmp',
            'image/gif'  => 'gif',
        ];
        $imageMime = image_type_to_mime_type($imageType);
        if (isset($extensions[$imageMime])) {
            $ctx = $this->getContext();
            $extension = $extensions[$imageMime];
            if (null === $ctx->contentTypes) {
                $ctx->contentTypes = new DOMDocument();
                $ctx->contentTypes->loadXML($ctx->tempDocumentContentTypes);
            }
            $xpath = new DOMXPath($ctx->contentTypes);
            $xpath->registerNamespace('ns', static::CONTENT_TYPES_NS);
            $matchedNode = null;
            $insertNode = null;
            /** @var \DOMElement $node */
            foreach ($xpath->query('//ns:Types/ns:Default[@Extension]') as $node) {
                if ($node->getAttribute('Extension') === 'rels') {
                    $insertNode = $node;
                }
                if ($node->getAttribute('Extension') === $extension) {
                    $matchedNode = $node;
                    break;
                }
            }
            if (null === $matchedNode) {
                $node = $ctx->contentTypes->createElementNS(static::CONTENT_TYPES_NS, 'Default');
                $node->setAttribute('Extension', $extension);
                $node->setAttribute('ContentType', $imageMime);
                if ($insertNode) {
                    $ctx->contentTypes->documentElement->insertBefore($node, $insertNode);
                } else {
                    $ctx->contentTypes->documentElement->appendChild($node);
                }
            }
            return $extension;
        }
    }

    /**
     * Get image relation id and update document relationships.
     *
     * @param string $extension Image extension
     * @param string $rid Original relation id from template
     * @return array
     */
    protected function getImageRelationId($extension, $rid)
    {
        $imageName = null;
        $ctx = $this->getContext();
        if (null == $ctx->relations) {
            $ctx->relations = new DOMDocument();
            $ctx->relations->loadXML($ctx->tempDocumentRelations[$ctx->getMainPartName()]);
        }
        $relId = in_array($rid, $ctx->relationIds) ? null : $rid;
        $usedIds = [];
        $xpath = new DOMXPath($ctx->relations);
        $xpath->registerNamespace('ns', static::RELATIONSHIPS_NS);
        /** @var \DOMElement $node */
        foreach ($xpath->query('//ns:Relationships/ns:Relationship') as $node) {
            $usedIds[] = (int) substr($node->getAttribute('Id'), 3);
            if (null !== $relId) {
                if ($node->getAttribute('Id') === $relId) {
                    /** @var \ZipArchive $zip */
                    $zip = $ctx->zipClass;
                    $zip->deleteName(sprintf('word/%s', $node->getAttribute('Target')));
                    $ctx->relations->documentElement->removeChild($node);
                    break;
                }
            }
        }
        if (null === $relId) {
            usort($usedIds, function($a, $b) {
                return $a - $b;
            });
            $i = 0;
            while (true) {
                $i++;
                if (!in_array($i, $usedIds)) {
                    $relId = sprintf('rId%d', $i);
                    break;
                }
            }
        }
        $ctx->relationIds[] = $relId;
        if (null === $imageName) {
            $imageName = sprintf('media/image%d.%s', count($ctx->relationIds), $extension);
        }
        $node = $ctx->relations->createElementNS(static::RELATIONSHIPS_NS, 'Relationship');
        $node->setAttribute('Id', $relId);
        $node->setAttribute('Type', static::IMAGE_RELATIONSHIP_TYPE);
        $node->setAttribute('Target', $imageName);
        $ctx->relations->documentElement->appendChild($node);
        return [$relId, $imageName];
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
            if (!$replace instanceof Symbol) {
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
     * @return array
     */
    protected function findTableTemplateRow($template)
    {
        $result = [$template, null];
        $startRow = null;
        $endRow = null;
        $matches = [];
        preg_match_all('#<w:tr(.*?)>(.*?)</w:tr>#', $template, $matches, PREG_OFFSET_CAPTURE);
        if ($this->extra) {
            $extras = explode(',', $this->extra);
            $startIdx = (int) $extras[0];
            $rowSize = count($extras) > 1 ? (int) $extras[1] : 1;
            $startIdx--;
            $endIdx = $startIdx + $rowSize - 1;
            if ($endIdx < count($matches[0])) {
                $startRow = $matches[0][$startIdx];
                $endRow = $matches[0][$endIdx];
            }
        } else {
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
        }
        if (null !== $startRow && null !== $endRow) {
            $result[0] = substr($template, $startRow[1], $endRow[1] + strlen($endRow[0]) - $startRow[1]);
            $result[1] = str_replace($result[0], '%ROWS%', $template);
        }
        return $result;
    }

    /**
     * Prepare build.
     *
     * @return \NTLAB\Report\Filer\DocumentTag
     */
    protected function prepareBuild()
    {
        $this->contents = [];
        $this->vars = $this->getVariables();
        $this->ifs = $this->findSubTags(static::DOC_SUB_IF, $this->vars);
        $this->eaches = $this->findSubTags(static::DOC_SUB_EACH, $this->vars);
        $this->tables = $this->findSubTags(static::DOC_SUB_TABLE, $this->vars);
        switch ($this->docType) {
            case static::DOC_DOCUMENT:
                break;
            case static::DOC_TABLE:
                list($template, $docTemplate) = $this->findTableTemplateRow($this->tempDocumentMainPart);
                $this->tempDocumentMainPart = $template;
                $this->docTemplate = $docTemplate;
                break;
            case static::DOC_EACH:
                break;
        }
        $this->images = $this->findImages($this->vars);
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
                if ($this->docTemplate) {
                    if (false !== strpos($this->docTemplate, '%ROWS%')) {
                        $this->tempDocumentMainPart = str_replace('%ROWS%', $this->tempDocumentMainPart, $this->docTemplate);
                    }
                }
                break;
            case static::DOC_EACH:
                break;
        }
        if ($this->contentTypes) {
            $this->tempDocumentContentTypes = $this->contentTypes->saveXML();
        }
        if ($this->relations) {
            $this->tempDocumentRelations[$this->getMainPartName()] = $this->relations->saveXML();
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