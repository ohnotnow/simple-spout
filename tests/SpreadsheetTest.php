<?php
// @codingStandardsIgnoreFile

use PHPUnit\Framework\TestCase;
use Ohffs\SimpleSpout\ExcelSheet;
use Carbon\Carbon;

class SpreadsheetTest extends TestCase
{
    /** @test */
    public function importing_a_spreadsheet_returns_correct_data()
    {
        $data = (new ExcelSheet)->import('./tests/data/test.xlsx');

        $row = $data[1];
        $this->assertEquals(Carbon::createFromFormat('d/m/Y H:i:s', '22/09/2015 00:00:00'), $row[0]);
        $this->assertEquals('5pm', $row[1]);
        $this->assertEquals('ENG4020', $row[2]);
    }

    /** @test */
    public function generating_a_sheet_from_an_array_produces_correct_results()
    {
        $spreadsheet = new ExcelSheet;
        $data = [
            0 => [
                'Jimmy', '01/02/2015', '1234567'
            ],
            1 => [
                'Fred', '02/03/2016', '9292929'
            ]
        ];

        $filename = $spreadsheet->generate($data);
        $newData = $spreadsheet->import($filename);
        unlink($filename);

        $this->assertEquals('Jimmy', $newData[0][0]);
        $this->assertEquals('01/02/2015', $newData[0][1]);
        $this->assertEquals('1234567', $newData[0][2]);
        $this->assertEquals('Fred', $newData[1][0]);
        $this->assertEquals('02/03/2016', $newData[1][1]);
        $this->assertEquals('9292929', $newData[1][2]);
    }
}
