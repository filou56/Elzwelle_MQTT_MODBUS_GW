copy ..\elzwelle_mqtt_gw.py .
mkdir dist
mkdir dist\elzwelle_mqtt_gw
\opt\miniconda3\Scripts\pyinstaller.exe elzwelle_mqtt_gw.py
copy \opt\miniconda3\Library\bin\libcrypto-3-x64.dll dist\elzwelle_mqtt_gw\_internal
copy \opt\miniconda3\Library\bin\libssh2.dll dist\elzwelle_mqtt_gw\_internal
