@echo off
setlocal enabledelayedexpansion

:: Cartella di input e output
set "input_dir=D:\P\_Audio trascrizione\opus"
set "output_dir=D:\P\_Audio trascrizione\My_audio"

:: Crea la cartella di output se non esiste
if not exist "%output_dir%" mkdir "%output_dir%"

:: Loop attraverso tutti i file .opus nella cartella di input
for %%f in ("%input_dir%\*.opus") do (
    :: Estrae il nome del file senza estensione
    set "filename=%%~nf"
    :: Converte il file .opus in .wav e lo salva nella cartella di output
    ffmpeg -i "%%f" "%output_dir%\!filename!.wav"
)

echo Conversione completata.
pause
