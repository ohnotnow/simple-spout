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

        $row = $data[0];
        $this->assertEquals('Cell A1', $row[0]);
        $this->assertEquals('Cell B1', $row[1]);
        $this->assertEquals('Cell C1', $row[2]);
        $row = $data[1];
        $this->assertEquals('Cell A2', $row[0]);
        $this->assertEquals('Cell B2', $row[1]);
        $this->assertEquals('Cell C2', $row[2]);
        $row = $data[2];
        $this->assertEquals('    Cell A3 With Extra Spaces.  ', $row[0]);
        $this->assertEquals('Cell B3', $row[1]);
        $this->assertEquals('Cell C3', $row[2]);
    }

    /** @test */
    public function we_can_import_just_the_first_sheet_from_an_xlsx_file()
    {
        $data = (new ExcelSheet)->importFirst('./tests/data/test.xlsx');

        $row = $data[0];
        $this->assertEquals('Cell A1', $row[0]);
        $this->assertEquals('Cell B1', $row[1]);
        $this->assertEquals('Cell C1', $row[2]);
        $row = $data[1];
        $this->assertEquals('Cell A2', $row[0]);
        $this->assertEquals('Cell B2', $row[1]);
        $this->assertEquals('Cell C2', $row[2]);
        $row = $data[2];
        $this->assertEquals('    Cell A3 With Extra Spaces.  ', $row[0]);
        $this->assertEquals('Cell B3', $row[1]);
        $this->assertEquals('Cell C3', $row[2]);
    }

    /** @test */
    public function we_can_import_a_specific_sheet_from_an_xlsx_file()
    {
        $data = (new ExcelSheet)->importSheet('./tests/data/test.xlsx', 0);

        $row = $data[0];
        $this->assertEquals('Cell A1', $row[0]);
        $this->assertEquals('Cell B1', $row[1]);
        $this->assertEquals('Cell C1', $row[2]);
        $row = $data[1];
        $this->assertEquals('Cell A2', $row[0]);
        $this->assertEquals('Cell B2', $row[1]);
        $this->assertEquals('Cell C2', $row[2]);
        $row = $data[2];
        $this->assertEquals('    Cell A3 With Extra Spaces.  ', $row[0]);
        $this->assertEquals('Cell B3', $row[1]);
        $this->assertEquals('Cell C3', $row[2]);
    }

    /** @test */
    public function can_trim_whitespace_from_all_cells_while_importing()
    {
        $data = (new ExcelSheet)->trimmedImport('./tests/data/test.xlsx');

        $row = $data[2];
        $this->assertEquals('Cell A3 With Extra Spaces.', $row[0]);
        $this->assertEquals('Cell B3', $row[1]);
        $this->assertEquals('Cell C3', $row[2]);
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
