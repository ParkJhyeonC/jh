@echo off
echo ===================================================
echo 정서행동특성검사 분석기 서버를 시작합니다...
echo ===================================================
echo.
echo 필요한 패키지를 설치하는 중입니다...
call npm install
echo.
echo 서버를 구동합니다. 브라우저에서 http://localhost:3000 으로 접속하세요.
call npm run dev
pause
