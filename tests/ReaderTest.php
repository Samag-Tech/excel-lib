<?php

use PHPUnit\Framework\TestCase;
use SamagTech\ExcelLib\Reader;

class ReaderTest extends TestCase {

    //------------------------------------------------------------------------------

    public function testBuildArray() {

        $this->reader = new Reader('./tests/export_test', 'export.xlsx');

        $data = $this->reader->toArray();

        $this->assertIsArray($data);
        $this->assertCount(1, $data);
    }

    //------------------------------------------------------------------------------

    public function testBuildArrayMultisheet() {

        $this->reader = new Reader('./tests/export_test', 'export_multisheet.xlsx');

        $data = $this->reader->toArray();

        $this->assertIsArray($data);
        $this->assertCount(2, $data);
    }

    //------------------------------------------------------------------------------

    public function testBuildArrayMultisheetWithNames() {

        $this->reader = new Reader('./tests/export_test', 'export_multisheet.xlsx');

        $data = $this->reader->setSheetNames(['Foglio 1', 'Foglio 2'])->toArray();

        $this->assertIsArray($data);
        $this->assertCount(2, $data);
        $this->assertArrayHasKey('Foglio_1', $data);
        $this->assertArrayHasKey('Foglio_2', $data);
    }

    //------------------------------------------------------------------------------

    public function testBuildObjet() {

        $this->reader = new Reader('./tests/export_test', 'export.xlsx');

        $data = $this->reader->toObject();

        $this->assertIsObject($data);
    }

    //------------------------------------------------------------------------------

    public function testBuildObjectMultisheetWithNames() {

        $this->reader = new Reader('./tests/export_test', 'export_multisheet.xlsx');

        $data = $this->reader->setSheetNames(['Foglio 1', 'Foglio 2'])->toObject();
        $this->assertIsObject($data);
        $this->assertObjectHasAttribute('Foglio_1', $data);
        $this->assertObjectHasAttribute('Foglio_2', $data);
    }

    //------------------------------------------------------------------------------
}