@echo off
pushd "%~dp0"
call clean.bat
copy original\sample?.pptx .
for %%f in (sample?.pptx) do python ..\replace_fonts.py %%f
popd
