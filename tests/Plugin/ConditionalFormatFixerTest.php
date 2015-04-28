<?php


namespace PHPExcelFixer\StyleFixer\Plugin;


class ConditionalFormatFixerTest extends BasePluginTest
{
    public function test_execute()
    {
        $output = $this->getMock('\ZipArchive');
        $template = $this->getMock('\ZipArchive');
        $bookUtil = $this->createBookUtil(2, ['Sheet1' => 'xl/worksheets/sheet1.xml', 'Sheet2' => 'xl/worksheets/sheet2.xml']);

        $output
            ->expects($this->exactly(2))
            ->method('getFromName')
            ->with($this->logicalOr('xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml'))
            ->will($this->returnCallback([$this, 'getOutputXml']))
        ;
        $template
            ->expects($this->once())
            ->method('getFromName')
            ->with($this->equalTo('xl/worksheets/sheet2.xml'))
            ->will($this->returnCallback([$this, 'getTemplateXml']))
        ;
        $output
            ->expects($this->once())
            ->method('addFromString')
            ->with($this->equalTo('xl/worksheets/sheet2.xml'), $this->callback([$this, 'assertXmlHasExpectedContent']))
        ;

        $fixer = new ConditionalFormatFixer($bookUtil);
        $fixer->execute($output, $template);
    }

    public function assertXmlHasExpectedContent($xml)
    {
        $this->assertContains('<conditionalFormatting><font/></conditionalFormatting>', $xml);
        $this->assertNotContains('<conditionalFormatting/>', $xml);

        return true;
    }
}
