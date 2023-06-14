<?php

namespace app\common\library;


class XlsWriter
{

    //服务器缓存地址
    const TEMP_PATH = RUNTIME_PATH.'excel';

    private $fileObject;
    private $excelName;


    /**
     * 创建一个excel
     * @param $paramsHeader
     * @param $paramsData
     * @param string $sheetName
     * @param string $excelName
     */
    public function createExcel($paramsHeader, $paramsData, $sheetName='Sheet1',$excelName= '')
    {
        $this->excelName = $excelName ? $excelName : '表格'.time();
        $tempPath = self::TEMP_PATH;
        if (!is_dir(dirname($tempPath))) {
            mkdir(dirname($tempPath), 0777, true);
        }
        $config = [
            'path' => $tempPath
        ];

        $excel = new \Vtiful\Kernel\Excel($config);

        $this->fileObject = $excel->constMemory("{$this->excelName}.xlsx",$sheetName);

        $fileHandle = $this->fileObject->getHandle();
        $format    = new \Vtiful\Kernel\Format($fileHandle);
        $boldStyle = $format->bold()->toResource();

        $this->fileObject->header($paramsHeader)
            ->data($paramsData);

    }


    /**
     * 向文件中追加一个工作表
     * @param $paramsHeader
     * @param $paramsData
     * @param string $sheetName
     */
    public function addExcel($paramsHeader, $paramsData, $sheetName='Sheet2')
    {
        $this->fileObject->addSheet($sheetName)
            ->header($paramsHeader)
            ->data($paramsData);
    }


    /**
     * 输出excel
     */
    public function outputAll()
    {
        $filePath = $this->fileObject->output();
        //@unlink($filePath);
    }
}
