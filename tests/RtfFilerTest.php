<?php

namespace NTLAB\Report\Test;

use NTLAB\RtfTree\Node\Tree;

class RtfFilerTest extends BaseTest
{
    public function testExtract()
    {
        $tree = new Tree();
        $this->assertTrue($tree->loadFromString($this->loadFixture('Template.rtf')), 'Template loaded succesdfully');
    }
}