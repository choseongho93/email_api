<?php

namespace App\Classes\Common;

use App\Constants\Common\EmailConstants;
use Illuminate\Support\Facades\Mail;

class Email
{

    private $email_view_data = EmailConstants::EMAIL_VIEW_DATA; // HTML view 데이터
    private $to; // 받는 사람
    private $subject; // 제목
    private $cc; // 참조
    private $attach_file; //첨부파일

    public function __construct($to)
    {
        $this->to = $to;
    }

    ## 기본 메일 발송
    public function basic_email()
    {

        Mail::send([], $this->email_view_data, function ($message) {
            $message->to($this->to)
                ->cc($this->cc)
                ->subject($this->subject);
        });
    }

    ## 파일첨부 메일 발송
    public function attachment_email()
    {
        Mail::send([], $this->email_view_data, function ($message) {
            $message->to($this->to)
                ->cc($this->cc)
                ->subject($this->subject)
                ->attach($this->attach_file);
        });
    }

    ## 제목 세팅
    public function setSubject($subject): Email
    {
        $this->subject = $subject;
        return $this;
    }

    ## 참조 세팅
    public function setCc($cc): Email
    {
        $this->cc = $cc;
        return $this;
    }

    ## 첨부파일 세팅
    public function setAttachFile($attach_file): Email
    {
        $this->attach_file = $attach_file;
        return $this;
    }

}