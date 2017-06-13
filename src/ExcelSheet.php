<?php

namespace Ohffs\SimpleSpout;

use Box\Spout\Reader\ReaderFactory;
use Box\Spout\Writer\WriterFactory;
use Box\Spout\Common\Type;

class ExcelSheet
{
    public function import($filename)
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
        return $rows;
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
}
