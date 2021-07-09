<?php


namespace App\Http\Controllers\Lunaplus;


use App\Classes\Common\Excel;
use App\Constants\ExcelConstants;
use App\Http\Controllers\Controller;
use App\Repositories\IEmailListRepo;
use Illuminate\Support\Facades\File;


class EmailController extends Controller
{
    private EmailListRepo;

    public function __construct
    (
        IEmailListRepo $EmailListRepo
    )
    {
        $this->EmailListRepo = $EmailListRepo;
    }

    ## 월 무료사용업체 엑셀파일로 자동발송
    public function EmailFreeEnd()
    {
        $now_ym = date("Y-m");

        $data = $this->EmailListRepo->getEmailFreeEndList($now_ym);

        foreach ($data as $item) {
            $array_data[] = [
                "user_id" => $item['user_id'],
                "login_id" => $item['login_id'],
                "tel_num" => $item['tel_num'],
            ];
        }

        $headers = ['유저', '로그인ID', '번호'];
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
        foreach ($array_data as $row) {
            $new_row_data[] = [
                $row['user_id'],
                $row['login_id'],
                $row['tel_num'],
            ];
        }
        $excel_data['header'][0] = $header_row_all;

        $tmp_path_name = $this->makeExcel($new_row_data, $excel_data, 'User');

        $this->sendEmailReminder($tmp_path_name);
        File::delete($tmp_path_name);

        exit;


    }


    private function makeExcel($new_row_data, $excel_data, $file_name)
    {
        $col_cnt = count($new_row_data[0]);

        $excel_data['data']['content'] = Collect($new_row_data)->map(function ($item, $key) {
            return Collect($item)->values();
        })->toArray();

        for ($i = 0; $i < $col_cnt; $i++) {
            $excel_data['data']['style'][] = ExcelConstants::DATA_BORDER_NUMBER_FORMAT;
        }

        $excel_data['filename'] = $this->makeFileName($file_name);

        $excel = new Excel();
        $excel->makeExcel($excel_data);

        $tmp_path_name = storage_path() ."/app/public/". $excel_data['filename'];
        $excel->saveXlsxExcel($tmp_path_name);

        return $tmp_path_name;
    }

    private function makeFileName($title)
    {
        list($usec, $sec) = explode(" ", microtime());
        $usec_str = round($usec * 1000000);
        $filename = sprintf('data_%s_%s_%06d_' . $title . '.xlsx', date('Ymd'), date('His'), $usec_str);
        return $filename;
    }

    public function sendEmailReminder($tmp_path_name)
    {

        $this->attachment_email($tmp_path_name);

    }


    public function basic_email() {
        $data = array('name'=>"Virat Gandhi");

        Mail::send([], $data, function($message) {
            $message->to('abc@gmail.com', 'Tutorials Point')->subject
            ('Laravel HTML Testing Mail');
            $message->from('xyz@gmail.com','리스트 정기 발송');
        });
        echo "Basic Email Sent. Check your inbox.";
    }

    public function html_email() {
        $data = array('name'=>"Virat Gandhi");
        Mail::send('mail', $data, function($message) {
            $message->to('abc@gmail.com', 'Tutorials Point')->subject
            ('Laravel HTML Testing Mail');
            $message->from('xyz@gmail.com','Virat Gandhi');
        });
        echo "HTML Email Sent. Check your inbox.";
    }

    public function attachment_email($tmp_path_name) {
        $data = array('name'=>"Virat Gandhi");

        Mail::send([], $data, function($message) use ($tmp_path_name) {
            $message->to('abc@gmail.com', 'Tutorials Point')->subject
            ('Laravel HTML Testing Mail');
            $message->attach($tmp_path_name);
            $message->from('xyz@gmail.com','리스트 정기 발송');
        });
        echo "Email Sent with attachment. Check your inbox.";
    }



}