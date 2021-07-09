<?php

namespace App\Classes\Common;

use App\Constants\ExcelConstants;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Writer\Xls;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class Excel
{
    var $objSpreadsSheet;

    /**
     *  생성자
     *  파일 기본 정보 작성, 디폴트 옵션 설정
     */
    public function __construct()
    {
        $this->objSpreadsSheet = new Spreadsheet();

        $properties = $this->objSpreadsSheet->getProperties();
        $properties->setCreator("nate");
        $properties->setLastModifiedBy("nate");
        $properties->setTitle("nate");
        $properties->setSubject("nate");

        $this->objSpreadsSheet->setActiveSheetIndex(0);
        $activeSheet = $this->objSpreadsSheet->getActiveSheet();

        // 디폴트 옵션 (폰트:맑은고딕, 크기:9, 셀 가로세로 : 가운데 정렬, 디폴트 높이 : 16.50)
        $defaultStyle = $this->objSpreadsSheet->getDefaultStyle();
        $defaultStyle->getFont()->setName('맑은 고딕');
        $defaultStyle->getFont()->setSize(9);
        $defaultStyle->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $defaultStyle->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);

        $activeSheet->getDefaultColumnDimension()->setWidth(11);
        $activeSheet->getDefaultRowDimension()->setRowHeight(16.50);
    }

    /**
     *  전반적 엑셀 파일 작성
     *  @return void
     */
    public function writeExcel($excel_data)
    {
        $this->makeExcel($excel_data);
        $this->downloadExcel($this->objSpreadsSheet, $excel_data['filename']);
    }

    /**
     *  전반적 엑셀 파일 작성
     *  @return void
     */
    public function makeExcel($excel_data)
    {
        $this->writeHeader($excel_data['header']);

        if (isset($excel_data['data'])) {
            $offset_data = count($excel_data['header']) + 1;

            $this->writeData($offset_data, $excel_data['data']);

            if (isset($excel_data['summary'])) {
                $offset_summary = count($excel_data['header']) + count($excel_data['data']['content']) + 1;

                $this->writeSummary($offset_summary, $excel_data['summary']);
            }
        }

        if (isset($excel_data['etc'])) {
            $this->writeEtc($excel_data['etc']);
        }
    }

    /**
     *  헤더 작성
     *  @return void
     */
    public function writeHeader($header_array)
    {
        $activeSheet = $this->objSpreadsSheet->getActiveSheet();

        foreach ($header_array as $key_row => $row_data) {

            $offset_col = ord('A');
            foreach ($row_data as $key_col => $col_data) {
                $colspan = isset($col_data['colspan']) ? $col_data['colspan'] : 0;
                $rowspan = isset($col_data['rowspan']) ? $col_data['rowspan'] : 0;

                $row = $key_row + 1;

                $offset_col = $this->getOffsetCol($row, $offset_col);
                $col = $this->chrExcelColumn($offset_col);

                // 셀 병합 여부 확인 후 셀을 병합하여 병합된 범위를 리턴
                $range = $this->getRange($col, $row, $colspan, $rowspan);

                // cell width
                if (isset($col_data['width'])) {
                    $activeSheet->getColumnDimension($col)->setWidth($col_data['width']);
                }

                // title
                if (isset($col_data['title'])) {
                    $activeSheet->setCellValue($col.$row, $col_data['title']);
                }

                // style
                if (isset($col_data['style'])) {
                    $activeSheet->getStyle($range)->applyFromArray($col_data['style']);
                }

                $offset_col++;
            }
        }
    }

    /**
     *  엑셀 데이터 작성
     *  @return void
     */
    public function writeData($offset, $data_array)
    {
        $activeSheet = $this->objSpreadsSheet->getActiveSheet();

        foreach ($data_array['content'] as $key_row => $data_row) {
            foreach ($data_row as $key_col => $data_col) {
                $col = $this->chrExcelColumn(ord('A')+$key_col);
                $row = $offset + $key_row;
                if (isset($data_array['style'][$key_col]['dataType']))
                    $activeSheet->setCellValueExplicit($col.$row, $data_col, $data_array['style'][$key_col]['dataType']);
                else
                    $activeSheet->SetCellValue($col.$row, $data_col);
            }
        }

        if (isset($data_array['style'])) {
            foreach ($data_array['style'] as $key => $style) {
                $col = $this->chrExcelColumn(ord('A')+$key);

                if (array_key_exists(0, $style)) {
                    foreach ($style as $index => $row) {
                        $range = $col.($offset + $index);
                        $activeSheet->getStyle($range)->applyFromArray($row);
                    }
                } else {
                    $range = $col . ($offset) . ":" . $col . ($offset + count($data_array['content']) - 1);

                    $activeSheet->getStyle($range)->applyFromArray($style);
                }
            }
        }
    }

    /**
     *  합계/요약 작성
     *  @return void
     */
    public function writeSummary($offset_row, $data_summary)
    {
        $activeSheet = $this->objSpreadsSheet->getActiveSheet();

        foreach ($data_summary as $key_row => $row_data) {

            $offset_col = ord('A');
            foreach ($row_data as $key_col => $col_data) {
                $colspan = isset($col_data['colspan']) ? $col_data['colspan'] : 0;
                $rowspan = isset($col_data['rowspan']) ? $col_data['rowspan'] : 0;

                $row = $offset_row + $key_row;

                $offset_col = $this->getOffsetCol($row, $offset_col);
                $col = $this->chrExcelColumn($offset_col);


                // 셀 병합 여부 확인 후 셀을 병합하여 병합된 범위를 리턴
                $range = $this->getRange($col, $row, $colspan, $rowspan);

                // title
                if (isset($col_data['title'])) {
                    $activeSheet->setCellValue($col.$row, $col_data['title']);
                }

                // style
                if (isset($col_data['style'])) {
                    $activeSheet->getStyle($range)->applyFromArray($col_data['style']);
                }

                $offset_col++;
            }
        }
    }

    /**
     *  기타 사항 작성 (코멘트 등)
     *  @return void
     */
    public function writeEtc($etc_array)
    {
        if (isset($etc_array['comment'])) {
            $this->writeComment($etc_array['comment']);
        }

        // 기타 작성 사항 들
    }

    /**
     *  기타 코멘트 작성
     *  @return void
     */
    public function writeComment($comment_array)
    {
        $activeSheet = $this->objSpreadsSheet->getActiveSheet();

        foreach ($comment_array as $comment) {
            $col = preg_replace("/[^A-Z]*/s", "", $comment['position']);
            $row = preg_replace("/[^0-9]*/s", "", $comment['position']);
            $colspan = isset($comment['colspan']) ? $comment['colspan'] : 0;
            $rowspan = isset($comment['rowspan']) ? $comment['rowspan'] : 0;

            // 셀 병합 여부 확인 후 셀을 병합하여 병합된 범위를 리턴
            $range = $this->getRange($col, $row, $colspan, $rowspan);

            // title
            if (isset($comment['content'])) {
                $activeSheet->setCellValue($comment['position'], $comment['content']);
            }

            // style
            if (isset($comment['style'])) {
                $activeSheet->getStyle($range)->applyFromArray($comment['style']);
            }
        }
    }

    /**
     *  데이터를 입력하려는 셀이 다른 셀에 의해 병합 되었는지 체크 후 입력 위치(컬럼) 리턴
     *  @return Integer
     */
    public function getOffsetCol($row, $offset)
    {
        $activeSheet = $this->objSpreadsSheet->getActiveSheet();

        $col = $this->chrExcelColumn($offset);

        $cell = $activeSheet->getCell($col.$row);

        // getMergeCells() : 통합된 셀의 범위들의 배열을 리턴
        // isInRange($cell) : 인자로 들어온 셀의 범위 안에 있는지 여부(boolean)를 리턴
        foreach ($activeSheet->getMergeCells() as $cells) {
            while ($cell->isInRange($cells)) {
                $offset++;
                $col = $this->chrExcelColumn($offset);

                $cell = $activeSheet->getCell($col.$row);
            }
        }

        return $offset;
    }

    /**
     *  셀을 병합하여 병합된 범위를 리턴
     *  @return String
     */
    public function getRange($col, $row, $colspan = 0, $rowspan = 0)
    {
        $activeSheet = $this->objSpreadsSheet->getActiveSheet();

        // 셀 병합
        if ($colspan && $rowspan) {
            $range = $col.$row.":".($this->chrExcelColumn2($col, $colspan - 1)).($row + ($rowspan - 1));
            $activeSheet->mergeCells($range);
        } elseif ($colspan) {
            $range = $col.$row.":".($this->chrExcelColumn2($col, $colspan - 1)).$row;
            $activeSheet->mergeCells($range);
        } elseif ($rowspan) {
            $range = $col.$row.":".$col.($row + ($rowspan - 1));
            $activeSheet->mergeCells($range);
        } else {
            $range = $col.$row;
        }

        return $range;
    }

    /**
     *  칼럼 개수가 26개를 초과하는 경우에 대한 처리
     *  @param String
     *  @return String
     */
    public function chrExcelColumn($col)
    {
        $chr = 'A';
        $cnt = $col - 65;
        for($i = 0; $i < $cnt; $i++) {
            $chr++;
        }
        return $chr;
    }
    public function chrExcelColumn2($col, $offset)
    {
        for($i = 0; $i < $offset; $i++) {
            $col++;
        }
        return $col;
    }

    /**
     *  엑셀 파일을 다운로드
     *  @return void
     */
    public function downloadExcel($objSpreadsSheet, $file_name)
    {
        header('Content-Type: application/vnd.ms-excel;');
        header('Content-Disposition: attachment; filename="'.$file_name.'.xls"');
        header('Cashe-Control: max-age=0');

        $writer = new Xls($objSpreadsSheet);
        $writer->save('php://output');
    }

    /**
     *  엑셀 파일을 로컬 경로에 저장
     *  @return void
     */
    public function saveExcel($path_file_name)
    {
        $writer = new Xls($this->objSpreadsSheet);
        $writer->save($path_file_name);
    }
    /**
     *  엑셀 파일을 로컬 경로에 저장
     *  @return void
     */
    public function saveXlsxExcel($path_file_name)
    {
        $writer = new Xlsx($this->objSpreadsSheet);
        $writer->save($path_file_name);
    }

    /**
     *  엑셀 파일 읽기
     *  @param $file, $maxColumn, $header
     *  @return Array
     */
    public function readExcel($file, $maxcol, $header = null)
    {
        $php_excel = IOFactory::load($file);

        $sheet = $php_excel->getSheet(0);                       // 첫번째 시트
        $maxRow = $sheet->getHighestRow();                      // 마지막 라인
        $maxColumn = $maxcol;                                   // 마지막 칼럼

        $read_target = "A1:$maxColumn"."$maxRow";
        $read_lines = $sheet->rangeToArray($read_target, '', TRUE, FALSE);

        // 파라미터로 양식 검사를 위한 헤더가 넘어오지 않은 경우에는 엑셀에서 읽은 배열을 그대로 반환
        if (! $header) return $read_lines;

        if ($read_lines[0] !== $header) return false;

        unset($read_lines[0]);

        return $read_lines;
    }


}
