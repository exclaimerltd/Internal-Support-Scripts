@ECHO OFF

Set dir=%appdata%\Microsoft\Signatures

Echo Deleting all files from %dir%
del %dir%\* /F /Q

Echo Deleting all folders from %dir%
for /d %%p in (%dir%\*) Do rd /Q /S "%%p"
@echo Folder deleted.


exit
