<?php
class Vtiful\Kernel\Excel {
/* 方法 */
public __construct(array $config)
public addSheet(string $sheetName)
public autoFilter(string $scope)
public constMemory(string $fileName, string $sheetName = ?)
public data(array $data)
public fileName(string $fileName, string $sheetName = ?)
public getHandle()
public header(array $headerData)
public insertFormula(int $row, int $column, string $formula)
public insertImage(int $row, int $column, string $localImagePath)
public insertText(
    int $row,
    int $column,
    int|float|string $data,
    string $format = ?
)
public mergeCells(string $scope, string $data)
public output()
public setColumn(string $range, float $width, resource $format = ?)
public setRow(string $range, float $height, resource $format = ?)
}
