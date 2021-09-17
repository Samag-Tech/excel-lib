<?php

use PHPUnit\Framework\TestCase;
use SamagTech\ExcelLib\AbstractExcel;

require 'Utils/PhpUnitUtils.php';

class AbstractExcelTest extends TestCase {

    private string $pathTestForCreate = './tests/create_for_test';

    protected function setUp() : void {
        $this->concrete = new class('./tests') extends AbstractExcel {};
    }

    //------------------------------------------------------------------------------

    public function testFilenameWithExt() {

        $this->concrete->setFilename('test.xlsx');

        $this->assertSame('test.xlsx', $this->concrete->getFilename());
    }

    //------------------------------------------------------------------------------

    public function testFilenameWithoutExt() {

        $this->concrete->setFilename('test');

        $this->assertSame('test.xlsx', $this->concrete->getFilename());
    }

    //------------------------------------------------------------------------------

    public function testPathWithSlash() {

        $this->concrete->setPath('test/');

        $this->assertSame('test/', $this->concrete->getPath());
    }

    //------------------------------------------------------------------------------

    public function testPathWithoutSlash() {

        $this->concrete->setPath('test');

        $this->assertSame('test/', $this->concrete->getPath());
    }

    //------------------------------------------------------------------------------

    public function testIgnoreFieldsFormat() {

        $array = ['test1', 'test2'];

        $this->concrete->setIgnoreFieldsFormat($array);

        $this->assertSame($array, $this->concrete->getIgnoreFieldsFormat());

    }

    //------------------------------------------------------------------------------

    /**
     * @depends testPathWithoutSlash
     *
     */
    public function testCreatePath () {

        $this->concrete->setPath($this->pathTestForCreate);

        $this->assertTrue(PhpUnitUtils::callMethod($this->concrete, 'createPath'));
        $this->assertDirectoryExists($this->pathTestForCreate);
        $this->assertDirectoryIsWritable($this->pathTestForCreate);

    }
}