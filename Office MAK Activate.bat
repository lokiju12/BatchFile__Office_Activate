@echo off
:: MS-OFFICE ��ǰ ���� ��ũ��Ʈ
:: �׽�Ʈ �غ� ������ 2016 2019 2021 �Դϴ�.

:: ���� ������ ���� Ű�� �ִٸ� �����ϴ� ����
echo ospp.vbs ��η� �̵� �� . . .
c:
cd C:\Program Files\Microsoft Office\Office16
echo ospp.vbs ��ɾ� ���� �� . . .
for /f "tokens=1,2 delims=:" %%a in ('cscript ospp.vbs /dstatus 2^>^&1 ^| find "product key"') do (
    set "product_key=%%b"
)
echo product_key �������� �յ� ���� ���� �� . . .
set "product_key=!product_key: =!"
set "product_key=!product_key:~0,5!"
echo ����� product Ű ���� ��� �� . . .
echo product key : %product_key%
echo Ű ���� ���� �� . . .
cscript ospp.vbs /unpkey:%product_key%

:: /inpkey: �ڿ� ��ǰ ����Ű�� �Է�
echo.
echo ���ͳݸ� OFFICE MAK ���̼��� Ű �Է� �� . . .
cscript ospp.vbs /inpkey:AAAAA-AAAAA-AAAAA-AAAAA

:: ���ǽ� ��ǰ ���� ����
echo OFFICE ��ǰ ���� Ȱ��ȭ �� . . .
cscript ospp.vbs /act
echo ���� ������ �Ϸ� �Ǿ����ϴ�.
echo ���α׷��� ���� �غ��ñ� �ٶ��ϴ�.
pause

