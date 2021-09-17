<?php

use PHPUnit\Framework\TestCase;
use SamagTech\ExcelLib\Reader;

class ReaderTest extends TestCase {

    //------------------------------------------------------------------------------

    protected function setUp(): void {
        $this->reader = new Reader('./tests/export_test', 'export_multisheet.xlsx');
    }

    //------------------------------------------------------------------------------

    public function testBuildArray() {
        $this->reader->buildArray();
    }

    //------------------------------------------------------------------------------

    public function testBuildArrayMultisheet() {
        $this->reader->buildArray();
    }

    //------------------------------------------------------------------------------
}