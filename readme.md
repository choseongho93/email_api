# email_api
laravel 프레임워크에서 email을 보내는 API를 구현
+ phpoffice를 통해 excel파일 첨부하여 메일 발송

## 사용 가이드
1. EmailController에서 엑셀파일로 만들 데이터를 Repository를 이용하여 데이터 조회
2. 데이터를 통해 엑셀파일로 만듬
3. sendEmail 함수를 호출해서 파일을 파라미터로 넘겨주고 메일 발송 (보내는 사람, 참조 등등 상수로 처리)
