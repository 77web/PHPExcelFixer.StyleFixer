<?php


namespace PHPExcel\StyleFixer\Plugin;


interface Plugin
{
    public function execute(\ZipArchive $output, \ZipArchive $template);
}
