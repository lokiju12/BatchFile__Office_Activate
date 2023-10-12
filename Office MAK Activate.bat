@echo off
:: MS-OFFICE 정품 인증 스크립트
:: 테스트 해본 버전은 2016 2019 2021 입니다.

:: 기존 인증에 사용된 키가 있다면 제거하는 과정
echo ospp.vbs 경로로 이동 중 . . .
c:
cd C:\Program Files\Microsoft Office\Office16
echo ospp.vbs 명령어 구동 중 . . .
for /f "tokens=1,2 delims=:" %%a in ('cscript ospp.vbs /dstatus 2^>^&1 ^| find "product key"') do (
    set "product_key=%%b"
)
echo product_key 변수에서 앞뒤 공백 제거 중 . . .
set "product_key=!product_key: =!"
set "product_key=!product_key:~0,5!"
echo 추출된 product 키 정보 출력 중 . . .
echo product key : %product_key%
echo 키 정보 삭제 중 . . .
cscript ospp.vbs /unpkey:%product_key%

:: /inpkey: 뒤에 정품 인증키를 입력
echo.
echo 인터넷망 OFFICE MAK 라이선스 키 입력 중 . . .
cscript ospp.vbs /inpkey:AAAAA-AAAAA-AAAAA-AAAAA

:: 오피스 정품 인증 시작
echo OFFICE 정품 인증 활성화 중 . . .
cscript ospp.vbs /act
echo 인증 과정이 완료 되었습니다.
echo 프로그램을 실행 해보시기 바랍니다.
pause

