@echo off
chcp 65001 >nul
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"
set "COMPOSE_PROJECT_NAME=documents_workspace"
echo 正在关闭 110 项目的 Docker 服务...
if not exist "runtime" mkdir "runtime" >nul 2>nul
>"runtime\.storage-migration.stop" echo stop
docker compose stop
if errorlevel 1 (
  echo 关闭失败，请确认 Docker Desktop 已启动并运行 docker compose ps 查看状态。
  pause
  exit /b 1
)
echo 已关闭 db、api、worker、web 服务。数据仍保留在配置的数据库和文件目录。
pause
endlocal
