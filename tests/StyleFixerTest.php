<?php


namespace PHPExcel\StyleFixer;


class StyleFixerTest extends \PHPUnit_Framework_TestCase
{
    private $basePath;

    public function setUp()
    {
        $this->basePath = __DIR__.'/xml';
    }

    public function getOutputXml($path)
    {
        return file_get_contents($this->basePath.'/output/'.$path);
    }

    public function getTemplateXml($path)
    {
        return file_get_contents($this->basePath.'/template/'.$path);
    }

    /**
     * @test
     */
    public function プラグインなし()
    {
        $output = $this->getMock('\ZipArchive');
        $template = $this->getMock('\ZipArchive');
        $outputPath = '/path/to/output.xlsx';
        $templatePath = '/path/to/template.xlsx';
        $zips = [
            $outputPath => $output,
            $templatePath => $template,
        ];

        $fixer = $this->getMockBuilder('PHPExcel\StyleFixer\StyleFixer')
            ->setMethods(['openFile'])
            ->getMock()
        ;
        $fixer
            ->expects($this->exactly(2))
            ->method('openFile')
            ->with($this->logicalOr($outputPath, $templatePath))
            ->will($this->returnCallback(function($path) use ($zips){
                        return $zips[$path];
                    }))
        ;
        $template
            ->expects($this->once())
            ->method('getFromName')
            ->with('xl/styles.xml')
            ->will($this->returnCallback([$this, 'getTemplateXml']))
        ;
        $output
            ->expects($this->once())
            ->method('addFromString')
            ->with('xl/styles.xml', $this->callback(function($xml){

                        $this->assertEquals($this->getTemplateXml('xl/styles.xml'), $xml);

                        return true;
                    }))
        ;
        $output
            ->expects($this->once())
            ->method('close')
        ;

        $fixer->execute($outputPath, $templatePath);
    }

    /**
     * @test
     */
    public function プラグインあり()
    {
        $output = $this->getMock('\ZipArchive');
        $template = $this->getMock('\ZipArchive');
        $outputPath = '/path/to/output.xlsx';
        $templatePath = '/path/to/template.xlsx';
        $zips = [
            $outputPath => $output,
            $templatePath => $template,
        ];
        $plugin = $this->getMockForAbstractClass('PHPExcel\StyleFixer\Plugin\Plugin');

        $fixer = $this->getMockBuilder('PHPExcel\StyleFixer\StyleFixer')
            ->setMethods(['openFile'])
            ->setConstructorArgs([[$plugin]])
            ->getMock()
        ;
        $fixer
            ->expects($this->exactly(3))
            ->method('openFile')
            ->with($this->logicalOr($outputPath, $templatePath))
            ->will($this->returnCallback(function($path) use ($zips){
                        return $zips[$path];
                    }))
        ;
        $plugin
            ->expects($this->once())
            ->method('execute')
            ->with($output, $template)
        ;
        $template
            ->expects($this->once())
            ->method('getFromName')
            ->with('xl/styles.xml')
            ->will($this->returnCallback([$this, 'getTemplateXml']))
        ;
        $output
            ->expects($this->once())
            ->method('addFromString')
            ->with('xl/styles.xml', $this->isType('string'))
        ;
        $output
            ->expects($this->exactly(2))
            ->method('close')
        ;

        $fixer->execute($outputPath, $templatePath);
    }

    public function tearDown()
    {
        $this->output_stylesXml = '';
        $this->template_stylesXml = '';
    }
}
