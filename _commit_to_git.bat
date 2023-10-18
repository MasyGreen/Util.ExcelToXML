set DD=%date:~0,2%
set MM=%date:~3,2%
set YYYY=%date:~6,4%
set Hour=%time:~0,2%
set Min=%time:~3,2%
set DT=_%YYYY%%MM%%DD%_%Hour%%Min%
SET DEBUGDATE=%DT%
SET PVERSION=Last

git add .
git commit -m "%PVERSION%"
git tag -a %PVERSION% -m "%PVERSION%"
git push origin master                                                        	
git push --tags                                                               	
git push origingit master                                                        	
pause