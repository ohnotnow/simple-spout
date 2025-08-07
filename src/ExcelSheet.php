<?php

namespace Ohffs\SimpleSpout;

use OpenSpout\Reader\XLSX\Reader;
use OpenSpout\Writer\XLSX\Writer;
use OpenSpout\Common\Entity\Row;

class ExcelSheet
{
    public function import($filename, $trim = false)
    {
        $reader = new Reader();
        $reader->open($filename);
        $rows = $this->getRowsFromSheets($reader);
        $reader->close();
        if ($trim) {
            $rows = $this->trim($rows);
        }
        return $rows;
    }

    public function importFirst($filename, $trim = false)
    {
        $reader = new Reader();
        $reader->open($filename);
        $rows = $this->getRowsFromSheets($reader, 0);
        $reader->close();
        if ($trim) {
            $rows = $this->trim($rows);
        }
        return $rows;
    }

    public function importSheet($filename, $sheetNumber = 0, $trim = false)
    {
        $reader = new Reader();
        $reader->open($filename);
        $rows = $this->getRowsFromSheets($reader, $sheetNumber);
        $reader->close();
        if ($trim) {
            $rows = $this->trim($rows);
        }
        return $rows;
    }

    public function importActive($filename, $trim = false)
    {
        $reader = new Reader();
        $reader->open($filename);
        $rows = $this->getRowsFromSheets($reader, $sheetNumber);
        $reader->close();
        if ($trim) {
            $rows = $this->trim($rows);
        }
        return $rows;
    }

    public function trimmedImport($filename)
    {
        return $this->import($filename, true);
    }

    public function generate($data, $filename = null)
    {
        if (!$filename) {
            $filename = tempnam('/tmp', 'sim');
        }
        $writer = new Writer();
        $writer->openToFile($filename);
        foreach ($data as $rowArray) {
            $writer->addRows([Row::fromValues($rowArray)]);
        }
        $writer->close();
        return $filename;
    }

    private function trim($rows)
    {
        $trimmedRows = [];
        foreach ($rows as $row) {
            $trimmedRows[] = array_map([$this, 'trimValue'], $row);
        }
        return $trimmedRows;
    }

    private static function trimValue($value)
    {
        if (! is_scalar($value)) {
            return $value;
        }
        return trim($value);
    }

    private function getRowsFromSheets($reader, $onlySheet = null, $onlyActive = false)
    {
        $rows = [];
        foreach ($reader->getSheetIterator() as $sheet) {
            if ($onlySheet !== null && $sheet->getIndex() !== $onlySheet) {
                continue;
            }
            if ($onlyActive && !$sheet->isActive()) {
                continue;
            }
            foreach ($sheet->getRowIterator() as $row) {
                $rows[] = $row->toArray();
            }
        }
        return $rows;
    }
}
