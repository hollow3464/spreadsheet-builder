<?php

namespace Hollow3464\SpreadsheetBuilder;

use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;

class Cell
{
    public function __construct(
        public int|string|null $content = null,
        public string $type = DataType::TYPE_STRING,
        public string $format = NumberFormat::FORMAT_GENERAL,
        public ?string $background = null,
        public ?string $textColor = null,
    ) {}
}
