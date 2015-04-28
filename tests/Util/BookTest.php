<?php


namespace PHPExcelFixer\StyleFixer\Util;

class BookTest extends \PHPUnit_Framework_TestCase
{
    public function test_makeSheetMap()
    {
        $zip = $this->makeZipMock(['xl/workbook.xml', 'xl/_rels/workbook.xml.rels']);

        $bookUtil = new Book();
        $map = $bookUtil->makeSheetMap($zip);

        $this->assertInternalType('array', $map);
        $this->assertArrayHasKey('表紙', $map);
        $this->assertArrayHasKey('日別_Y', $map);
        $this->assertArrayHasKey('日別_YDN', $map);
        $this->assertArrayNotHasKey('日別_G', $map);
        $this->assertEquals('xl/worksheets/sheet1.xml', $map['表紙']);
    }

    public function test_makePrintAreaMap()
    {
        $zip = $this->makeZipMock(['xl/workbook.xml']);
        $bookUtil = new Book();
        $map = $bookUtil->makePrintAreaMap($zip);

        $this->assertInternalType('array', $map);
        $this->assertArrayHasKey('表紙', $map);
        $this->assertArrayHasKey('日別_Y', $map);
        $this->assertArrayHasKey('日別_YDN', $map);
        $this->assertArrayNotHasKey('日別_G', $map);
        $this->assertEquals('\'日別_Y\'!$A$18:$A$18', $map['日別_Y']);
    }

    public function test_makeSheetRelationMap()
    {
        $zip = $this->makeZipMock(['xl/workbook.xml', 'xl/_rels/workbook.xml.rels']);

        $bookUtil = new Book();
        $map = $bookUtil->makeSheetRelationMap($zip);

        $this->assertInternalType('array', $map);
        $this->assertArrayHasKey('表紙', $map);
        $this->assertArrayHasKey('日別_Y', $map);
        $this->assertArrayHasKey('日別_YDN', $map);
        $this->assertArrayNotHasKey('日別_G', $map);
        $this->assertEquals('xl/worksheets/_rels/sheet1.xml.rels', $map['表紙']);
    }

    public function test_makeDrawingMap()
    {
        $zip = $this->makeZipMock(['xl/workbook.xml', 'xl/_rels/workbook.xml.rels', 'xl/worksheets/_rels/sheet1.xml.rels', 'xl/worksheets/_rels/sheet2.xml.rels', 'xl/worksheets/_rels/sheet3.xml.rels']);

        $bookUtil = new Book();
        $map = $bookUtil->makeDrawingMap($zip);

        $this->assertInternalType('array', $map);
        $this->assertArrayNotHasKey('表紙', $map);
        $this->assertArrayHasKey('日別_Y', $map);
        $this->assertArrayHasKey('日別_YDN', $map);
        $this->assertEquals(['xl/drawings/drawing1.xml', 'xl/drawings/drawing2.xml'], $map['日別_Y']);
    }

    /**
     * @param string $fileNames
     * @return \ZipArchive|\PHPUnit_Framework_MockObject_MockObject
     */
    private function makeZipMock($fileNames)
    {
        $zip = $this->getMock('\ZipArchive');
        $zip
            ->expects($this->exactly(count($fileNames)))
            ->method('getFromName')
            ->with($this->callback(
                function($fileName) use ($fileNames){
                        return in_array($fileName, $fileNames);
                    }))
            ->will($this->returnCallback([$this, 'getXml']))
        ;

        return $zip;
    }

    public function getXml($path)
    {
        return file_get_contents(__DIR__.'/../xml/template/'.$path);
    }
}
