# Spreadsheet Builder

This is a small wrapper around the powerful
[PhpSpreadsheet](https://github.com/PHPOffice/PhpSpreadsheet)
package that allows to quickly build spreadsheets using a fluent builder.

## Usage

Building a sheet is more like traversing a matrix instead of the "setup this cell"
approach of the underlying package. We begin at the first position "A1", "0,0"
in the matrix and use the builder functions to move and setup the content for
the cells in the "Active Worksheet".

Instead of a pair axis we have the concept of columns and rows, so the "move" functions
use these to describe themselves as "move on <axis>" be it on the X axis (Columns)
or the Y axis (Rows).

```php
  $sheet = (new SpreadsheetBuilder())
      ->set('hello')
      ->moveColumn()
      ->set(new Cell('world'))
      ->build();

  // Output
  // A1: hello ; A2: world
```

These move functions can take an int as an argument to spectify how many spaces
will you move, but assume by default that you will move to the left and below.
The move functions also accept a negative value for travesing in backwards motion.

To reduce the amount of functions used, abbreviations are provided.

```php
    ->moveRowStart() // Move to the start of the row (Moves on column direction)
    ->moveColumnStart() // Move to the start of the column (Moves on row direction)
    ->setMoveColumn('hello')
    ->setMoveRow('world')
```

## Origin

I felt that the original package lacks a proper and straightforward solution
to quickly iterating over a spreadsheet and setting up common cell options
such as the type and formatting.

Requiring to always specify cell coordinates to access the underlying
cell model means that to fill out a sheet programatically either a little state
machine must be created to track the current cell position and such state be
modified per operation (set, move, merge blocks) for each project that intends
to use the package or a template must be stored and accesed during the generation
of the sheets. The latter can become cumbersome due to constraints around the development
process.

This package has the intent of becoming that thin state-machine layer allowing
a lot of flexibility and an easy to remember mental model around the sheet.
