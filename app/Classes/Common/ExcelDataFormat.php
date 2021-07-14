<?php

namespace App\Classes\Common;

use App\Constants\Common\ExcelConstants;

class ExcelDataFormat extends Excel
{
    private $new_row_data;
    private $excel_data;

    ## 엑셀 파일 데이터 가공
    public function setData($data, $key_val, $key_name)
    {
        $array_data = [];

        for ($i = 0; $i < count($data); $i++) {
            $temp_data = [];
            for ($j = 0; $j < count($key_val); $j++) {
                $column_key = $key_val[$j];

                $temp_data[$column_key] = $data[$i]->$column_key;
            }
            array_push($array_data, $temp_data);
        }

        $headers = $key_name;
        $header_row_all = [];
        foreach ($headers as $header) {
            $header_row_all[] = [
                'title' => $header,
                'width' => 20,
                'style' => ExcelConstants::HEADER_BORDER,
            ];
        }

        $excel_data = [];
        $new_row_data = [];

        for ($i = 0; $i < count($array_data); $i++) {
            $temp_data = [];
            for ($j = 0; $j < count($key_val); $j++) {
                $column_key = $key_val[$j];

                array_push($temp_data, $array_data[$i][$column_key]);
            }
            array_push($new_row_data, $temp_data);
        }

        $excel_data['header'][0] = $header_row_all;
        $this->new_row_data = $new_row_data;
        $this->excel_data = $excel_data;

        return $this;

    }

    ## 엑셀파일 만드는 함수
    public function makeExcelByFoarmat($file_name, $file_path)
    {
        $col_cnt = count($this->new_row_data[0]);

        $this->excel_data['data']['content'] = Collect($this->new_row_data)->map(function ($item) {
            return Collect($item)->values();
        })->toArray();

        for ($i = 0; $i < $col_cnt; $i++) {
            $this->excel_data['data']['style'][] = ExcelConstants::DATA_BORDER_NUMBER_FORMAT;
        }

        $this->excel_data['filename'] = $this->makeFileName($file_name);

        $this->makeExcel($this->excel_data);
        $tmp_path_name = $file_path . $this->excel_data['filename'];
        $this->saveXlsxExcel($tmp_path_name);

        return $tmp_path_name;
    }

    ## 파일 생성
    private function makeFileName($title)
    {
        list($usec, $sec) = explode(" ", microtime());
        $usec_str = round($usec * 1000000);
        $filename = sprintf('data_%s_%s_%06d_' . $title . '.xlsx', date('Ymd'), date('His'), $usec_str);

        return $filename;
    }
}