@echo off
:: ����Ƿ����Թ���Ա�������
openfiles >nul 2>nul
if %errorlevel% NEQ 0 (
    echo ��Ҫ����ԱȨ�ޡ������Թ���ԱȨ����������...
    powershell -Command "Start-Process cmd -ArgumentList '/c %~s0' -Verb RunAs"
    exit
)

:: ɾ����ǰ�ļ��������� .xlsx �ļ�
echo ����ɾ�� .xlsx �ļ�...
del /q *.xlsx
echo .xlsx �ļ���ɾ����

:: ��ѹ���� .zip ѹ��������ǰ�ļ���
echo ���ڽ�ѹ .zip ѹ����...
for %%i in (*.zip) do (
    echo ��ѹ %%i...
    powershell -Command "Expand-Archive -Path '%%i' -DestinationPath '.'"
)
echo ��ѹ��ɡ�

:: �������Զ��˳�
exit
