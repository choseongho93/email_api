<?php


namespace App\Http\Controllers\Lunaplus;


use App\Classes\Common\Email;
use App\Classes\Common\Excel;
use App\Classes\Common\ExcelDataFormat;
use App\Constants\Common\EmailConstants;
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
        try {
            ## 무료사용 만료업체 데이터 가져오기
            $data = $this->getFreeEndData();

            if(!empty($data)){
                ## 엑셀
                $key_val = ["c_name", "user_id", "login_id", "name"];
                $key_name = ['업체명', '대표운영자ID', '부운영자ID', '운영담당자'];

                ## 엑셀
                $excel = new ExcelDataFormat();
                $file_path = storage_path() . "/app/public/"; // 파일 저장 경로

                ## 엑셀파일 만들기
                $file = $excel->setData($data, $key_val, $key_name)->makeExcelByFoarmat('free_end', $file_path);

                if (!empty($file)) {
                    $this->sendEmail($file); // 메일 발송
                    File::delete($file); // 파일 삭제
                }
            }

            return $this->returnData(null, "메일 발송 완료");

        } catch (\Throwable $e) {
            return $this->returnFailed($e->getMessage());
        }

    }


    ## Email 발송 함수
    private function sendEmail($file)
    {
        $to_email = EmailConstants::OPERATION_TEAM_EMAIL;
        $email = new Email($to_email);
        $email->setSubject($this->date. "월 상담원 무료 만료 리스트 공유")
            ->setCc(EmailConstants::EMAIL_TEAM_CC)
            ->setAttachFile($file)
            ->attachment_email();
    }


}