# PHPExcel.StyleFixer
[![Build Status](https://travis-ci.org/77web/PHPExcelFixer.StyleFixer.svg)](https://travis-ci.org/77web/PHPExcelFixer.StyleFixer)

Fixes cell styles broken through [PHPExcel](https://github.com/phpoffice/phpexcel) read & write process.

## Installation

Use composer.

```
composer install phpexcel-fixer/style-fixer
```

## Usage

```php
<?php

use PHPExcelFixer\StyleFixer\StyleFixer;
use PHPExcelFixer\StyleFixer\Plugin\CellStyleFixer;
use PHPExcelFixer\StyleFixer\Plugin\ConditionalFormatFixer;

$templatePath = '/path/to/template.xlsx';
$outputPath = '/path/to/output.xlsx';

$fixer = new StyleFixer([new CellStyleFixer(), new ConditionlFormatFixer()]);
$fixer->execute($outputPath, $templatePath);
```

## Feedback

If you find any issue, please let me know.
https://github.com/77web/PHPExcelFixer.StyleFixer/issues
