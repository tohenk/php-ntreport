<?php

namespace NTLAB\Report\Test;

use NTLAB\Report\Util\RtfFiler;
use NTLAB\RtfTree\Node\Tree;

class RtfFilerTest extends BaseTest
{
    public function testExtract()
    {
        $tree = new Tree();
        $this->assertTrue($tree->loadFromString($this->loadFixture('Template.rtf')), 'Template loaded succesdfully');

        $body = RtfFiler::getInstance()->extract($tree, $bpos, $epos);
        $this->assertNotNull($body, 'Body extracted successfully');
        $this->saveOut($body->toStringEx(), 'body.txt');
    }
}