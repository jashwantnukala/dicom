@echo off
REM ============================================================================
REM Build script for dicom_receiver_pusher-7-4.py -> DicomReceiverPusher.exe
REM
REM Run this from a Windows command prompt, with dicom_receiver_pusher.spec,
REM dicom_receiver_pusher-7-4.py, requirements.txt, and (optionally)
REM lucide_icons_data.py / logo.png all in this same folder.
REM ============================================================================

echo Installing dependencies...
pip install -r requirements.txt
pip install pyinstaller

echo.
echo Building (this uses dicom_receiver_pusher.spec, which bundles the
echo DICOM codec plugins correctly -- see the comments at the top of that
echo file for why a plain "pyinstaller dicom_receiver_pusher-7-4.py" is not
echo enough)...
pyinstaller --noconfirm dicom_receiver_pusher.spec

echo.
echo Done. Build output is in dist\DicomReceiverPusher\
echo Launch dist\DicomReceiverPusher\DicomReceiverPusher.exe to test.
echo.
echo To verify the codec fix took effect, check that these folders exist:
dir /b "dist\DicomReceiverPusher\_internal" | findstr /i "dist-info" | findstr /i "pylibjpeg gdcm"
pause
