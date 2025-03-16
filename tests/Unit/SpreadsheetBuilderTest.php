<?php

use Hollow3464\SpreadsheetBuilder\Cell;
use Hollow3464\SpreadsheetBuilder\SpreadsheetBuilder;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

describe(SpreadsheetBuilder::class, function () {
    it('builds a spreadsheet', function () {
        $builder = new SpreadsheetBuilder;
        $spreadsheet = $builder->build();

        expect($spreadsheet)->toBeInstanceOf(Spreadsheet::class);
    });

    it('sets a value to a cell', function () {
        $builder = new SpreadsheetBuilder;
        $sheet = $builder->set('hello')->build();

        expect(
            $sheet
                ->getActiveSheet()
                ->getCell('A1')
                ->getValue()
        )->toBe('hello');
    });

    it('sets a value to a cell from a cell builder', function () {
        $builder = new SpreadsheetBuilder;

        $builder->set(new Cell('hello'));

        $sheet = $builder->build();

        expect(
            $sheet
                ->getActiveSheet()
                ->getCell('A1')
                ->getValue()
        )->toBe('hello');
    });

    it('moves to the next cell column', function () {
        $builder = new SpreadsheetBuilder;
        $builder->moveColumn();

        expect($builder->currentCell)->toBe('B1');
    });

    it('moves a column backwards', function () {
        $builder = new SpreadsheetBuilder;
        $builder->moveColumn(2);
        $builder->moveColumn(-1);

        expect($builder->currentCell)->toBe('B1');
    });

    it('moves to the next row', function () {
        $builder = new SpreadsheetBuilder;
        $builder->moveRow();

        expect($builder->currentCell)->toBe('A2');
    });

    it('moves a row backwards', function () {
        $builder = new SpreadsheetBuilder;
        $builder->moveRow(2);
        $builder->moveRow(-1);

        expect($builder->currentCell)->toBe('A2');
    });

    it('moves to the first column in the row', function () {
        $builder = new SpreadsheetBuilder;

        $builder
            ->moveColumn()
            ->moveRowStart();

        expect($builder->currentCell)->toBe('A1');
    });

    it('moves to the first row in the column', function () {
        $builder = new SpreadsheetBuilder;

        $builder
            ->moveRow()
            ->moveColumnStart();

        expect($builder->currentCell)->toBe('A1');
    });

    it('sets a value and moves to the next cell', function () {
        $builder = new SpreadsheetBuilder;
        $builder->setMoveRow('hello');

        expect($builder->sheet->getActiveSheet()->getCell('A1')->getValue())->toBe('hello');
        expect($builder->currentCell)->toBe('B1');
    });

    it('sets a value and moves to the cell below', function () {
        $builder = new SpreadsheetBuilder;
        $builder->setMoveColumn('hello');

        expect($builder->sheet->getActiveSheet()->getCell('A1')->getValue())->toBe('hello');
        expect($builder->currentCell)->toBe('A2');
    });

    it('sets content to a merge block and moves to the next cell at the start', function () {
        $builder = new SpreadsheetBuilder;
        $builder->setMergeBlock('hello', 2, 2);

        $mergeCells = $builder->active->getMergeCells();

        expect($mergeCells)->not()->toBeEmpty();
        expect(array_keys($mergeCells)[0])->toBe('A1:B2');
        expect(
            $builder->sheet
                ->getActiveSheet()
                ->getCell('A1')->getValue()
        )->toBe('hello');

        expect($builder->currentCell)->toBe('C1');
    });

    it('sets content to a merge block and moves to the next cell at the end', function () {
        $builder = new SpreadsheetBuilder;
        $builder->setMergeBlock('hello', 2, 2, alignment: 'end');

        expect($builder->active->getMergeCells())
            ->not()
            ->toBeEmpty();

        expect($builder->sheet->getActiveSheet()->getCell('A1')->getValue())->toBe('hello');
        expect($builder->currentCell)->toBe('C2');
    });

    it('sets content to a merge block and moves to the cell below at the start', function () {
        $builder = new SpreadsheetBuilder;
        $builder->setMergeBlock('hello', 2, 2, 'down');

        expect($builder->active->getMergeCells())
            ->not()
            ->toBeEmpty();

        expect($builder->sheet->getActiveSheet()->getCell('A1')->getValue())->toBe('hello');
        expect($builder->currentCell)->toBe('A3');
    });

    it('sets content to a merge block and moves to the cell below at the end', function () {
        $builder = new SpreadsheetBuilder;
        $builder->setMergeBlock('hello', 2, 2, 'down', 'end');

        expect($builder->active->getMergeCells())
            ->not()
            ->toBeEmpty();

        expect($builder->sheet->getActiveSheet()->getCell('A1')->getValue())->toBe('hello');
        expect($builder->currentCell)->toBe('B3');
    });

})->coversClass(SpreadsheetBuilder::class);
