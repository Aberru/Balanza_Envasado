@echo off
echo ========================================
echo    GENERADOR DE EJECUTABLE BALANZA
echo ========================================
echo.

echo Instalando dependencias...
pip install -r requirements.txt

echo.
echo Generando ejecutable...
pyinstaller --clean BalanzaDataAcquisition.spec

echo.
echo ========================================
echo    EJECUTABLE GENERADO EXITOSAMENTE
echo ========================================
echo.
echo El ejecutable se encuentra en: dist\BalanzaDataAcquisition.exe
echo.
pause

