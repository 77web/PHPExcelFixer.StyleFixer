<?php


namespace PHPExcelFixer\StyleFixer\Plugin;


interface Plugin
{
    public function execute(\ZipArchive $output, \ZipArchive $template);
}
