<?php


namespace PHPExcelFixer\StyleFixer\Plugin;


abstract class BasePluginTest extends \PHPUnit_Framework_TestCase
{
    private $basePath;

    public function setUp()
    {
        $this->basePath = __DIR__.'/../xml';
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
     * @param int $count
     * @param array $map
     * @return \PHPUnit_Framework_MockObject_MockObject
     */
    protected function createBookUtil($count, $map)
    {
        $bookUtil = $this->getMock('PHPExcelFixer\StyleFixer\Util\Book');
        $bookUtil
            ->expects($this->exactly($count))
            ->method('makeSheetMap')
            ->with($this->isInstanceOf('\ZipArchive'))
            ->will($this->returnValue($map))
        ;

        return $bookUtil;
    }

    public function tearDown()
    {
        $this->output_stylesXml = '';
        $this->template_stylesXml = '';
    }
}
