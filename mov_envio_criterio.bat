@echo off
REM Mover arquivos Excel da área de trabalho para a pasta destino
move "%UserProfile%\Desktop\criterios\*.xlsx" ..\..\..\projetospython\envio_criterio\app\data\

REM Navegar para o diretório do projeto e rodar o script Python com Poetry
cd ..\..\..\projetospython\envio_criterio\app
poetry run python main.py

REM Esperar 15 segundos antes de encerrar
timeout /t 15
