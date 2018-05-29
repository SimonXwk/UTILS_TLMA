@echo off
set env_folder=env

if exist "%env_folder%\" (
	echo ^>^>^> found virtual environment [%env_folder%]
) else (
	echo ^>^>^> virtual environment [%env_folder%] can not be found :^(
	rd /s /q "env"
	echo ^>^>^> creating virtual environment [%env_folder%], it might take a while ... please wait 
	python -m venv env
)

echo ^>^>^> activate virtual environment ^& collect python dependecies
cmd /k "env\Scripts\activate & pip install -r requirements.txt -U & echo Ready :^)"