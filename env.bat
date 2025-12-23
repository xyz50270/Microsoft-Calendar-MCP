@echo off
set USE_MODELSCOPE_HUB=1
@set CUDA_VISIBLE_DEVICES=0
rem SET FORCE_TORCHRUN=1
rem SET USE_LIBUV=0
if "%VIRTUAL_ENV%"=="%cd%\.venv" (
	goto end
) else (
    if not exist ".venv" (
		echo 正在创建新的虚拟环境
		python -m venv .venv
		@CALL .venv\Scripts\activate.bat
		python -m pip install --upgrade pip
		pip install -U wheel
		@cmd /k "%*"
		@CALL .venv\Scripts\deactivate.bat
	) else (
		@CALL .venv\Scripts\activate.bat
		@cmd /k "%*"
		@CALL .venv\Scripts\deactivate.bat
	)
)
:end