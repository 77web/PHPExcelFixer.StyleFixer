<?php


namespace PHPExcel\StyleFixer\Util;


class SheetTest extends \PHPUnit_Framework_TestCase
{
    /**
     * @param string $cell
     * @param string $range
     * @param bool $expect
     * @test
     * @dataProvider provideInRangeTestData
     */
    public function test_inRange($cell, $range, $expect)
    {
        $sheetUtil = new Sheet();
        $this->assertEquals($expect, $sheetUtil->inRange($cell, $range));
    }

    public function provideInRangeTestData()
    {
        return [
            ['A1',  'A2:A10', false],
            ['A2',  'A2:A10', true],
            ['A5',  'A2:A10', true],
            ['A10', 'A2:A10', true],
            ['A11', 'A2:A10', false],
            ['A1', 'B1:Z1', false],
            ['B1', 'B1:Z1', true],
            ['C1', 'B1:Z1', true],
            ['Z1', 'B1:Z1', true],
            ['AA1', 'B1:Z1', false],
            ['A1', 'AA1:AA2', false],
        ];
    }
}
