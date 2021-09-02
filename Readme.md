# Basic PHP spreadsheet import/export

This is just a basic wrapper around 'Spout' (https://github.com/box/spout). It will read a spreadsheet
and convert it to a PHP array, and can take a PHP array and convert it to an excel file.

## Installation

Just do a `composer require ohffs/simple-spout`.

## Usage

```
<?php

namespace App\Whatever;

use Ohffs\SimpleSpout\ExcelSheet;

class Thing
{
    public function something()
    {
	    // plain import to array
        $data = (new ExcelSheet)->import('/tmp/spreadsheet.xlsx');
        ...;
	    // if you want each cell to have whitespace trimmed from the beginning/end
	    $data = (new ExcelSheet)->trimmedImport('/tmp/spreadsheet.xlsx');

        // import just the very first sheet
        $data = (new ExcelSheet)->importFirst('/tmp/spreadsheet.xlsx');

        // import a specific sheet
        $data = (new ExcelSheet)->importSheet('/tmp/spreadsheet.xlsx', 3); // 0-indexed

        // import the 'active' sheet (ie, the one that was open when the file was saved)
        $data = (new ExcelSheet)->importActive('/tmp/spreadsheet.xlsx');
    }

    public function somethingElse()
    {
        $data = [
          ['smith', 'sarah-jane', 'companion'],
          ['baker', 'tom', 'doctor'],
        ];
        $filename = (new ExcelSheet)->generate($data);
        // or
        (new ExcelSheet)->generate($data, '/data/spreadsheet.xlsx');
    }
}
```
