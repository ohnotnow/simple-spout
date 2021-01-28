# Basic PHP spreadsheet import/export

This is just a basic wrapper around 'Spout' (https://github.com/box/spout). It will read a spreadsheet
and convert it to a PHP array, and can take a PHP array and convert it to an excel file.

## Installation

Add to your composer.json file entries for :

```
{
    "require": {
        "ohffs/simple-spout": "^1.0"
    }
}
```

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
