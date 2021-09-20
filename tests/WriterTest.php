<?php

use PHPUnit\Framework\TestCase;
use SamagTech\ExcelLib\ExcelException;
use SamagTech\ExcelLib\Writer;

class WriterTest extends TestCase {

    private string $pathTestBuild = './tests/export_test';
    private Writer $writer;

    //------------------------------------------------------------------------------

    protected function setUp(): void {
        $this->writer = new Writer($this->pathTestBuild);
    }

    //------------------------------------------------------------------------------

    /**
     * @depends AbstractExcelTest::testFilenameWithoutExt
     * @depends AbstractExcelTest::testPathWithoutSlash
     *
     */
    public function testBuild() {

        $headers = ['A'];

        $body = [
            ['A1']
        ];

        $this->writer->setFilename('export');

        $excelPath = $this->writer->setHeader($headers)
            ->setBody($body)
            ->build();

        $this->assertSame($excelPath, $this->writer->getPath().$this->writer->getFilename());
        $this->assertDirectoryExists($this->pathTestBuild);
        $this->assertFileExists($this->pathTestBuild.'/export.xlsx');
    }

    //------------------------------------------------------------------------------

    /**
     * @depends AbstractExcelTest::testFilenameWithoutExt
     * @depends AbstractExcelTest::testPathWithoutSlash
     *
     */
    public function testBuildMultisheet() {

        $this->writer->setFilename('export_multisheet');

        $headers = [
            [
                'H1',
                'H2',
                'H3'
            ],
            [
                'h1',
                'h2'
            ]
        ];

        $body = [
            [
                [
                    'R1',
                    'R2',
                    'R3'
                ],
            ],
            [
                [
                    -1,
                    'r2'
                ]
            ]
        ];

        $excelPath = $this->writer->setHeader($headers)->setBody($body)->build(true);
        $this->assertSame($excelPath, $this->writer->getPath().$this->writer->getFilename());
        $this->assertDirectoryExists($this->pathTestBuild);
        $this->assertFileExists($this->pathTestBuild.'/export_multisheet.xlsx');

    }

    //------------------------------------------------------------------------------

     /**
     * @depends AbstractExcelTest::testFilenameWithoutExt
     * @depends AbstractExcelTest::testPathWithoutSlash
     *
     */
    public function testBuildMultisheetWithSingleHeader() {

        $this->writer->setFilename('export_multisheet_single_header');

        $headers = [
            'H1',
            'H2',
            'H3'
        ];

        $body = [
            [
                [
                    'R1',
                    'R2',
                    'R3'
                ],
            ],
            [
                [
                    'r1',
                    'r2'
                ]
            ]
        ];

        $excelPath = $this->writer->setHeader($headers)->setBody($body)->build(true);
        $this->assertSame($excelPath, $this->writer->getPath().$this->writer->getFilename());
        $this->assertDirectoryExists($this->pathTestBuild);
        $this->assertFileExists($this->pathTestBuild.'/export_multisheet.xlsx');

    }

    //------------------------------------------------------------------------------

    public function testBuildWithoutBody() {

        $this->expectException(ExcelException::class);
        $this->expectExceptionMessage('Il corpo dell\'excel è vuoto');
        $this->writer->build();
    }

    //------------------------------------------------------------------------------

    public function testBuildMalformedBody() {

        $this->expectException(ExcelException::class);
        $this->expectExceptionMessage('Il corpo dell\'excel è malformato');

        $this->writer->setBody(['test'])->build();

    }

    //------------------------------------------------------------------------------
}