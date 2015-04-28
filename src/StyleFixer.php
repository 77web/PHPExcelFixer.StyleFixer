<?php

namespace PHPExcelFixer\StyleFixer;

use PHPExcelFixer\StyleFixer\Plugin\Plugin;

/**
 * Class StyleFixer
 * PHPExcelが壊した書式設定を修復（テンプレートからコピー）する
 */
class StyleFixer
{
    const STYLES_XML_PATH = 'xl/styles.xml';

    /**
     * @var Plugin[]
     */
    private $plugins;

    public function __construct(array $plugins = null)
    {
        $this->plugins = $plugins;
    }
    
    /**
     * @param string $outputPath
     * @param string $templatePath
     */
    public function execute($outputPath, $templatePath)
    {
        $output = $this->openFile($outputPath);
        $template = $this->openFile($templatePath);

        // スタイル定義を修復
        $this->fixStyles($output, $template);
        $output->close();

        // 個別のシートを修復
        if (null !== $this->plugins) {
            foreach ($this->plugins as $fixer) {

                $output = $this->openFile($outputPath);

                $fixer->execute($output, $template);

                $output->close();
            }
        }

        $template->close();
    }

    /**
     * @param string $path
     * @return \ZipArchive
     */
    protected function openFile($path)
    {
        $zip = new \ZipArchive;
        $zip->open($path);
        
        return $zip;
    }

    /**
     * xl/styles.xml全体を修復（テンプレートからコピー）する
     *
     * @param \ZipArchive $output
     * @param \ZipArchive $template
     */
    private function fixStyles(\ZipArchive $output, \ZipArchive $template)
    {
        $srcStylesXml = $template->getFromName(self::STYLES_XML_PATH);

        $output->addFromString(self::STYLES_XML_PATH, $srcStylesXml);
    }
}
