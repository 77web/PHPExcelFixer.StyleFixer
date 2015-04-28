<?php


namespace PHPExcelFixer\StyleFixer\Util;


class Sheet
{

    /**
     * セル番地がExcel形式のセル範囲(コロン区切りのみ対応)に含まれるかどうか
     *
     * @param string $coordinate
     * @param string $cellRange
     * @return bool
     */
    public function inRange($coordinate, $cellRange)
    {
        // 絶対参照の$を削除
        $cellRange = str_replace('$', '', $cellRange);
        list($startCell, $endCell) = explode(':', $cellRange);

        // 行の範囲チェック
        $rowNumber = $this->detectRowNumberFromCoordinate($coordinate);
        if ($rowNumber < $this->detectRowNumberFromCoordinate($startCell) || $rowNumber > $this->detectRowNumberFromCoordinate($endCell)) {
            return false;
        }

        // 列の範囲チェック
        $colNumber = $this->detectColNumberFromCoordinate($coordinate);
        if ($colNumber < $this->detectColNumberFromCoordinate($startCell) || $colNumber > $this->detectColNumberFromCoordinate($endCell)) {
            return false;
        }

        return true;
    }

    /**
     * セル番地から行番号を取り出す
     *
     * @param string $coordinate
     * @return int
     */
    private function detectRowNumberFromCoordinate($coordinate)
    {
        return intval(preg_replace('/[A-Z]+/', '', $coordinate));
    }

    /**
     * セル番地から列番号を取り出す
     *
     * @param string $coordinate
     * @return int
     */
    private function detectColNumberFromCoordinate($coordinate)
    {
        $char = preg_replace('/\d+/', '', $coordinate);

        $charA = 65;

        if (strlen($char) == 1) {
            return ord($char) - $charA + 1;
        }

        $quot = ord($char[0]) - $charA + 1;
        $mod = ord($char[1]) - $charA + 1;

        return $quot * 26 + $mod;
    }
}
