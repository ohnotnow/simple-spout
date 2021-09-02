<?php

namespace Ohffs\SimpleSpout;

use Box\Spout\Reader\ReaderFactory;
use Box\Spout\Writer\WriterFactory;
use Box\Spout\Common\Type;
use Box\Spout\Reader\Common\Creator\ReaderEntityFactory;
use Box\Spout\Writer\Common\Creator\WriterEntityFactory;

class ExcelSheet
{
    public function import($filename, $trim = false)
    {
        $reader = ReaderEntityFactory::createXLSXReader();
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
        $reader = ReaderEntityFactory::createXLSXReader();
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
        $reader = ReaderEntityFactory::createXLSXReader();
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
        $reader = ReaderEntityFactory::createXLSXReader();
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
        $writer = WriterEntityFactory::createXLSXWriter();
        $writer->openToFile($filename);
        foreach ($data as $rowArray) {
            $writer->addRows([WriterEntityFactory::createRowFromArray($rowArray)]);
        }
        $writer->close();
        return $filename;
    }

    private function trim($rows)
    {
        $trimmedRows = [];
        foreach ($rows as $row) {
            $trimmedRows[] = array_map('trim', $row);
        }
        return $trimmedRows;
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
