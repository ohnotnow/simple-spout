<?php

namespace Ohffs\SimpleSpout;

use Box\Spout\Reader\ReaderFactory;
use Box\Spout\Writer\WriterFactory;
use Box\Spout\Common\Type;

class ExcelSheet
{
    public function import($filename, $trim = false)
    {
        $reader = ReaderFactory::create(Type::XLSX);
        $reader->open($filename);
        $rows = [];
        foreach ($reader->getSheetIterator() as $sheet) {
            foreach ($sheet->getRowIterator() as $row) {
                $rows[] = $row;
            }
        }
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
        $writer = WriterFactory::create(Type::XLSX);
        $writer->openToFile($filename);
        $writer->addRows($data);
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
}
