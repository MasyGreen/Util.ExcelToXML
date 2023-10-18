rem Удалить локальные изменения, синхронизировать с удаленным (EMART)
@ECHO =====================Attention!!!============================
@ECHO Your Solution synk from Git. All not commited change will be lost.
@ECHO =====================Attention!!!============================
pause
git fetch origin
git reset --hard origin/master
git clean -f -d