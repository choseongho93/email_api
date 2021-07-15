<?php


namespace App\Http\Controllers;


use App\Classes\Common\Email;
use App\Classes\Common\ExcelDataFormat;
use App\Constants\EmailConstants;
use App\Repositories\IEmailRepo;
use App\Traits\ResultFunction;
use Illuminate\Support\Facades\File;


class EmailController extends Controller
{
    use ResultFunction;

    private $EmailRepo;

    public function __construct
    (
        IEmailRepo $EmailRepo
    )
    {
        $this->EmailRepo = $EmailRepo;
    }


    public function EmailFreeEnd()
    {
        try {
            $data = $this->EmailRepo->getUser();

            if(!empty($data)){
                ## 엑셀
                $key_val = ["name", "login_id", "created_date"];
                $key_name = ['이름', '로그인 ID', '등록 날짜'];

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

            return $this->returnData(null, "success");

        } catch (\Throwable $e) {
            return $this->returnFailed($e->getMessage());
        }

    }


    ## Email 발송 함수
    private function sendEmail($file)
    {
        $to_email = EmailConstants::TO_TEAM_EMAIL;
        $email = new Email($to_email);
        $email->setSubject("subject")
            ->setCc(EmailConstants::EMAIL_TEAM_CC)
            ->setAttachFile($file)
            ->attachment_email();
    }


}