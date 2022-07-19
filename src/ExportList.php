<?php

declare(strict_types=1);

namespace Cxx\ExcelHelper;

use Cxx\ExcelHelper\ExcelException;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Writer\Xls;

/**
 * 数据导出
 */
class ExportList
{
    /**
     * 对应的表头
     *
     * @var array
     */
    protected $header = [];

    /**
     * 要导出的数据
     *
     * @var array
     */
    protected $exportData = [];

    /**
     * 导出模板
     *
     * @var string
     */
    protected $export_tpl;

    /**
     * 写入的开始行数
     *
     * @var integer
     */
    protected $start_write_line = 2;

    /**
     * @param array $header [
     *  '表头字段名称' => '对应数据字段',
     *  '姓名' => 'name'
     * ]
     * @param array $exportData
     */
    public function __construct(array $header, array $exportData)
    {
        $this->header = $header;
        $this->exportData = $exportData;
    }

    /**
     * 导出
     *
     * @param string $name 文件名(不需要文件后缀)
     * @return void
     */
    public function export(string $name = '')
    {
        $spreadsheet = $this->getSpreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        if (!$this->export_tpl) {
            $this->setHeader($sheet);
        }
        $this->writeData($sheet);
        $writer = new Xls($spreadsheet);
        // $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');
        $name = $name ?: date("YmdHis");
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="' . $name . '.xls"');
        header('Cache-Control: max-age=0');
        // If you're serving to IE 9, then the following may be needed
        header('Cache-Control: max-age=1');
        // If you're serving to IE over SSL, then the following may be needed
        header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
        header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header('Pragma: public'); // HTTP/1.0
        header('Access-Control-Allow-Origin: *');
        header('Access-Control-Allow-Methods: GET, POST, PATCH, PUT, DELETE, OPTIONS');
        header('Access-Control-Expose-Headers: Content-Disposition');
        header('Access-Control-Allow-Headers: Authorization, Content-Type, If-Match, If-Modified-Since, If-None-Match, If-Unmodified-Since, X-CSRF-TOKEN, X-Requested-With, responseType');
        $writer->save('php://output');
        exit();
    }

    /**
     * 获取表格实例
     *
     * @return Spreadsheet
     */
    public function getSpreadsheet(): Spreadsheet
    {
        if ($this->export_tpl) {
            return IOFactory::load($this->export_tpl);
        }
        return new Spreadsheet();
    }

    /**
     * 写入表头
     *
     * @param Worksheet $sheet
     * @return void
     */
    private function setHeader($sheet)
    {
        $a = $this->wordInc(count($this->header) > 26 ? 2 : 1);
        $i = 0;
        foreach ($this->header as $head => $field) {
            $sheet->setCellValue($a[$i++] . '1', $head);
        }
    }

    /**
     * 写入数据
     *
     * @param Worksheet $sheet
     * @return void
     */
    private function writeData($sheet)
    {
        $line = $this->start_write_line;
        $a = $this->wordInc(count($this->header) > 26 ? 2 : 1);
        $i = 0;
        foreach ($this->exportData as $item) {
            foreach ($this->header as $field) {
                $value = $item[$field] ?? '';
                $sheet->setCellValueExplicit(
                    $a[$i++] . $line,
                    $value,
                    DataType::TYPE_STRING
                );
            }
            $i = 0;
            $line++;
        }
    }

    /**
     * 单词自增序列 (生成的条目数为 $level * 26 )
     *
     * @param integer $level
     * @return array
     */
    public function wordInc(int $level = 1): array
    {
        $words = range('A', 'Z');
        if ($level <= 1) {
            return $words;
        }
        $level--;
        for ($j = 0; $j < $level; $j++) {
            for ($i = 0; $i < 26; $i++) {
                $words[] = $words[$j] . $words[$i];
            }
        }
        return $words;
    }

    /**
     * 设置导入模板
     * 注意：模板的表头顺序必须要跟 $header 一样
     *
     * @param string $export_tpl
     * @return $this
     */
    public function setExportTpl(string $export_tpl)
    {
        $this->export_tpl = $export_tpl;
        return $this;
    }

    /**
     * 设置开始写入行数
     * 
     * @param integer $start_write_line
     * @return $this
     */
    public function setStartWriteLine(int $start_write_line)
    {
        if ($start_write_line <= 1) {
            throw new ExcelException('start_write_line min 2');
        }
        $this->start_write_line = $start_write_line;
        return $this;
    }
}
