<?php

use SamagTech\ExcelLib\Reader;
use SamagTech\ExcelLib\Writer;
use PHPUnit\Framework\TestCase;
use SamagTech\ExcelLib\ExcelEnum;
use SamagTech\ExcelLib\ExcelException;

require 'Utils/MyConfig.php';
require 'Utils/MyFormatCell.php';

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

    /**
     * @depends ReaderTest::testBuildArray
     *
     */
    public function testBuildDateFormat() {

        $headers = ['A', 'B', 'C'];

        $body = [
            ['A' => '2021-01-01'],
            ['B' => '01-2021-01'],
            ['C' => '01/01/2021'],
        ];


        $this->writer->setFilename('export_with_format_date');
        $this->writer->setColumnDefinition([
            'A' => [
                'type'          => ExcelEnum::DEFINITION_DATE,
                'old_format'    => 'Y-m-d',
                'new_format'    => 'd/m/Y',
            ],
            'B' => [
                'type'          => ExcelEnum::DEFINITION_DATE,
                'old_format'    => 'm-Y-d',
                'new_format'    => 'd/m/Y',
            ],
            'C' => [
                'type'          => ExcelEnum::DEFINITION_DATE,
                'old_format'    => 'd/m/Y',
                'new_format'    => 'Y-m-d',
            ]
        ]);


        $excelPath = $this->writer->setHeader($headers)
            ->setBody($body)
            ->build();

        $this->reader = new Reader('./tests/export_test', 'export_with_format_date.xlsx');

        $data = $this->reader->toArray();

        $checkDate = function ($format, $date) : bool {
            $dt = DateTime::createFromFormat($format, $date);
            return $dt && $dt->format($format) === $date;
        };

        $this->assertTrue($checkDate('d/m/Y',$data[1][0]));
        $this->assertTrue($checkDate('d/m/Y',$data[2][0]));
        $this->assertTrue($checkDate('Y-m-d',$data[3][0]));
    }

    //------------------------------------------------------------------------------

    public function test_my_custom_config_and_format() {

        $writer = new Writer($this->pathTestBuild, 'my_custom', config: MyConfig::class, formatCell: MyFormatCell::class);

        $headers = ['A'];

        $body = [
            ['A1']
        ];

        $excelPath = $writer->setHeader($headers)
            ->setBody($body)
            ->build();

        $this->assertSame($excelPath, $writer->getPath().$writer->getFilename());
        $this->assertFileExists($this->pathTestBuild.'/my_custom.xlsx');

    }

    //------------------------------------------------------------------------------
}