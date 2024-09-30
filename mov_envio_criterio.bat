@echo off
REM Mover arquivos Excel da área de trabalho para a pasta destino
move "%UserProfile%\Desktop\criterios\*.xlsx" ..\..\..\projetospython\envio_criterio\app\data\

REM Navegar para o diretório do projeto e rodar o script Python com Poetry
cd ..\..\..\projetospython\envio_criterio\app
poetry run python main.py

REM Esperar 300 segundos antes de encerrar
timeout /t 300

REM Mover arquivos Excel de volta para area de trabalho
move "..\..\..\projetospython\envio_criterio\app\data\*.xlsx" %UserProfile%\Desktop\criterios\

del /q \projetospython\envio_criterio\app\jsons\*.*
