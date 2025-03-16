<?php

namespace Hollow3464\SpreadsheetBuilder;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class SpreadsheetBuilder
{
    public string $currentSheet = '';
    public string $currentCell = 'A1';
    public string $currentColumn = 'A';
    public int $currentRow = 1;
    public Spreadsheet $sheet;
    public Worksheet $active;

    public function __construct()
    {
        $this->sheet = $this->create();
        $this->active = $this->sheet->getActiveSheet();
    }

    public function create(): Spreadsheet
    {
        return new Spreadsheet;
    }

    public function build(): Spreadsheet
    {
        return $this->sheet;
    }

    public function nextColumn(string $coordinate, int $count = 1): string
    {
        return Coordinate::stringFromColumnIndex(
            Coordinate::columnIndexFromString($coordinate) + $count
        );
    }

    public function moveColumnStart(): self
    {
        $this->currentRow = 1;
        $this->currentCell = sprintf('%s%d', $this->currentColumn, $this->currentRow);

        return $this;
    }

    public function moveRowStart(): self
    {
        $this->currentColumn = 'A';
        $this->currentCell = sprintf('%s%d', $this->currentColumn, $this->currentRow);

        return $this;
    }

    public function moveColumn(int $moves = 1): self
    {
        $this->currentColumn = $this->nextColumn($this->currentColumn, $moves);
        $this->currentCell = sprintf('%s%d', $this->currentColumn, $this->currentRow);

        return $this;
    }

    public function moveRow(int $moves = 1): self
    {
        $this->currentRow = $this->currentRow + $moves;
        $this->currentCell = sprintf('%s%d', $this->currentColumn, $this->currentRow);

        return $this;
    }

    public function set(Cell|string|null $cell = ''): self
    {
        $activeCell = $this->active->getCell($this->currentCell);

        if (! is_object($cell)) {
            $activeCell->setValue($cell);

            return $this;
        }

        $activeCell->setValue($cell->content);

        if ($cell->type) {
            $activeCell->setDataType($cell->type);
        }

        if ($cell->format) {
            $activeCell
                ->getStyle()
                ->getNumberFormat()
                ->setFormatCode($cell->format);
        }

        return $this;
    }

    public function setMoveRow(Cell|string|null $content = '', int $moves = 1): self
    {
        $this->set($content);
        $this->moveColumn($moves);

        return $this;
    }

    public function setMoveColumn(Cell|string|null $content = ''): self
    {
        $this->set($content);
        $this->moveRow();

        return $this;
    }

    public function setMergeBlock(
        Cell|string|null $content = null,
        int $width = 1,
        int $height = 1,
        string $continue = 'after',
        string $alignment = 'start'
    ): self {
        $this->set($content);

        // Merge Logic
        $start = $this->currentCell;

        $startCoordinate = Coordinate::coordinateFromString($start);
        $startColumn = Coordinate::columnIndexFromString($startCoordinate[0]);
        $startRow = $startCoordinate[1];

        $endColumn = $startColumn + $width - 1;
        $endRow = $startRow + $height - 1;

        $end = Coordinate::stringFromColumnIndex($endColumn) . $endRow;

        $this->active->mergeCells("{$start}:{$end}");

        if ($continue === 'after') {
            $this->moveColumn($width);
            if ($alignment === 'end') {
                $this->moveRow($height - 1);
            }
        }

        if ($continue === 'down') {
            $this->moveRow($height);
            if ($alignment === 'end') {
                $this->moveColumn($width - 1);
            }
        }

        return $this;
    }

    /**
     * @param  array<array<int|string>>  $table
     */
    public function addTable(array $table): self
    {
        if (! count($table)) {
            return $this;
        }

        $rowLength = count($table[0]);

        foreach ($table as $row) {
            foreach ($row as $cell) {
                // Set and move to the next cell
                $this->setMoveRow($cell);
            }

            $this
                // Move to the next row
                ->moveRow()
                // Move to the first column of the table
                ->moveColumn(-$rowLength);
        }

        return $this;
    }

    public function getCurrentCell(): Cell
    {
        return $this->active->getCell($this->currentCell);
    }
}
