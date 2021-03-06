<?php

use PHPUnit\Framework\TestCase;
use SamagTech\ExcelLib\Factory;
use SamagTech\ExcelLib\Reader;
use SamagTech\ExcelLib\Writer;

class ExcelFactoryTest extends TestCase {

    protected function setUp(): void {
        $this->factory = new Factory();
    }

    //------------------------------------------------------------------------------
    public function testCreateWriter () {
        $this->assertTrue($this->factory->createWriter('test') instanceof Writer);
    }

    //------------------------------------------------------------------------------

    public function testCreateReader() {
        $this->assertTrue($this->factory->createReader('test', 'file.xlsx') instanceof Reader);
    }

    //------------------------------------------------------------------------------

}