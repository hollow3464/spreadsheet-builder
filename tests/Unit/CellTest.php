<?php

use Hollow3464\SpreadsheetBuilder\Cell;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat\Wizard\Currency;

describe(Cell::class, function () {
    it('allows PhpSpreadsheet types', function () {
        expect((new Cell(type: DataType::TYPE_NUMERIC))->type)
            ->toBe(DataType::TYPE_NUMERIC);
    });

    it('allows PhpSpreadshete masks and turns them into strings', function () {
        $currencyMask = new Currency;
        expect((new Cell(format: $currencyMask))->type)
            ->toBeString();
    });
})->covers(Cell::class);
