@echo off
chcp 65001 >nul
if /i "%~1"=="--migration-watcher" goto migration_watcher
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"
set "COMPOSE_PROJECT_NAME=documents_workspace"
title 110 Platform Startup
color 0F

set "STORAGE_HOST_PATH=./runtime/storage"
set "POSTGRES_HOST_PATH=./runtime/postgres"
if exist ".env" for /f "usebackq tokens=1,* delims== eol=#" %%A in (".env") do (
  if "%%A"=="STORAGE_HOST_PATH" set "STORAGE_HOST_PATH=%%B"
  if "%%A"=="POSTGRES_HOST_PATH" set "POSTGRES_HOST_PATH=%%B"
)
set "STORAGE_HOST_PATH=%STORAGE_HOST_PATH:"=%"
set "POSTGRES_HOST_PATH=%POSTGRES_HOST_PATH:"=%"

echo.
echo ==================================================
echo   110 File Management and Knowledge Search
echo   Startup check - Docker, Compose and all services
echo ==================================================
echo.

echo [1/8] Checking project files and host runtime...
if not exist "docker-compose.yml" (
  echo       ERROR: docker-compose.yml was not found in %cd%.
  echo       Run this BAT from the project root or restore the file.
  goto failed
)
if not exist "backend\Dockerfile" (
  echo       ERROR: backend\Dockerfile is missing.
  echo       Restore the backend directory before starting.
  goto failed
)
if not exist "frontend\Dockerfile" (
  echo       ERROR: frontend\Dockerfile is missing.
  echo       Restore the frontend directory before starting.
  goto failed
)
if not exist ".env" (
  echo       WARNING: .env is missing; compose defaults will be used.
  echo       If this is a copied project, review .env.example for optional settings.
) else (
  echo       .env: found
)
if not exist ".env.example" (
  echo       WARNING: .env.example is missing; optional environment template unavailable.
)
if not exist "%POSTGRES_HOST_PATH%" mkdir "%POSTGRES_HOST_PATH%" >nul 2>nul
if not exist "%STORAGE_HOST_PATH%" mkdir "%STORAGE_HOST_PATH%" >nul 2>nul
echo ok>"%STORAGE_HOST_PATH%\.startup_write_test"
if not exist "%STORAGE_HOST_PATH%\.startup_write_test" (
  echo       ERROR: storage directory is not writable: %STORAGE_HOST_PATH%
  echo       Create the directory and grant Docker/current user write permission.
  goto failed
)
del /q "%STORAGE_HOST_PATH%\.startup_write_test" >nul 2>nul
where curl.exe >nul 2>nul
if errorlevel 1 (
  echo       ERROR: curl.exe is unavailable; API and website health checks cannot run.
  echo       Install/enable Windows curl, then run this BAT again.
  goto failed
)
echo       Project files: OK
echo       Database path: %POSTGRES_HOST_PATH%
echo       Storage path: %STORAGE_HOST_PATH%
echo       Data directories: OK
echo       curl.exe: OK

echo.
echo [2/8] Checking Docker command and Compose...
where docker >nul 2>nul
if errorlevel 1 goto docker_missing
docker compose version >nul 2>nul
if errorlevel 1 goto compose_missing
echo       Docker command: OK
echo       Docker Compose: OK

if not exist "runtime" mkdir "runtime" >nul 2>nul
docker version --format "{{.Client.Version}}" >"runtime\.docker-client-version.tmp" 2>nul
if errorlevel 1 (
  echo       ERROR: Docker client version could not be read.
  echo       Repair Docker Desktop or add docker.exe to PATH.
  goto failed
)
set /p DOCKER_CLIENT_VERSION=<"runtime\.docker-client-version.tmp"
del /q "runtime\.docker-client-version.tmp" >nul 2>nul
echo       Docker client version: !DOCKER_CLIENT_VERSION!

docker info >nul 2>nul
if errorlevel 1 goto docker_not_ready
echo       Docker engine:  OK
goto docker_ready

:docker_missing
echo       ERROR: Docker command was not found.
echo       Install Docker Desktop and make sure docker.exe is in PATH.
goto failed

:compose_missing
echo       ERROR: Docker Compose is unavailable.
echo       Update Docker Desktop, then run: docker compose version
goto failed

:docker_not_ready
echo       Docker engine is not ready.
echo       [1] Start Docker Desktop automatically
echo       [2] Wait for me to start Docker Desktop manually
choice /c 12 /n /m "Select [1/2]: "
if errorlevel 2 goto wait_docker
set "DOCKER_EXE=%ProgramFiles%\Docker\Docker\Docker Desktop.exe"
if not exist "%DOCKER_EXE%" set "DOCKER_EXE=%LocalAppData%\Programs\DockerDesktop\Docker Desktop.exe"
if not exist "%DOCKER_EXE%" (
  echo       Docker Desktop executable was not found; waiting instead.
  goto wait_docker
)
echo       Starting Docker Desktop...
start "" "%DOCKER_EXE%"

:wait_docker
for /l %%i in (1,1,45) do (
  docker info >nul 2>nul
  if not errorlevel 1 goto docker_ready
  echo       Waiting for Docker engine... %%i/45
  timeout /t 2 /nobreak >nul
)
echo       Docker engine did not become ready in time.
echo       Start Docker Desktop, wait until it says Running, then run this file again.
goto failed

:docker_ready
echo.
echo [3/8] Checking Compose configuration...
docker compose config --quiet
if errorlevel 1 (
  echo       ERROR: docker-compose.yml is invalid or incompatible.
  echo       Run: docker compose config
  goto failed
)
echo       Compose configuration: OK
for %%S in (db api worker web) do (
  set "SERVICE_FOUND="
  for /f "delims=" %%L in ('docker compose config --services 2^>nul') do if /i "%%L"=="%%S" set "SERVICE_FOUND=1"
  if not defined SERVICE_FOUND (
    echo       ERROR: required Compose service %%S is missing.
    echo       Check docker-compose.yml and restore the complete project folder.
    goto failed
  )
)
echo       Required services: db, api, worker, web

echo.
echo [4/8] Checking db, api, worker and web services...
set "MISSING="
for %%S in (db api worker web) do (
  docker compose ps --services --filter status=running | findstr /c:"%%S" >nul 2>nul
  if errorlevel 1 (
    echo       %%S: missing or stopped
    set "MISSING=1"
  ) else (
    echo       %%S: running
  )
)
if not defined MISSING goto services_ready

echo.
echo [5/8] Starting missing services (build if needed)...
docker compose up -d --build
if errorlevel 1 goto diagnose
echo       Compose start command completed.

:services_ready
echo.
echo [6/8] Verifying all services and the API...
set "SERVICES_OK="
for /l %%i in (1,1,30) do (
  set "MISSING="
  for %%S in (db api worker web) do (
    docker compose ps --services --filter status=running | findstr /c:"%%S" >nul 2>nul
    if errorlevel 1 set "MISSING=1"
  )
  if not defined MISSING set "SERVICES_OK=1"
  if defined SERVICES_OK goto services_verified
  echo       Waiting for all services... %%i/30
  timeout /t 2 /nobreak >nul
)
goto diagnose

:services_verified
for %%S in (db api worker web) do echo       %%S: ready

set "API_OK="
for /l %%i in (1,1,30) do (
  curl.exe -fsS http://localhost:8000/api/health >nul 2>nul
  if not errorlevel 1 set "API_OK=1"
  if defined API_OK goto api_ready
  echo       Waiting for API health... %%i/30
  timeout /t 2 /nobreak >nul
)

:api_ready
if not defined API_OK goto diagnose
curl.exe -fsS http://localhost:5173/ >nul 2>nul
if errorlevel 1 goto diagnose

echo.
echo [7/8] Startup complete.
echo       API:     http://localhost:8000
echo       Website: http://localhost:5173
echo       Database data: %POSTGRES_HOST_PATH%
echo       File data:     %STORAGE_HOST_PATH%
call :start_migration_watcher
echo.
echo Opening the website in your default browser in 3 seconds...
timeout /t 3 /nobreak
start "" "http://localhost:5173"
echo.
echo ==================================================
echo  Startup finished. This window will stay open.
echo  Close it with any key. Services will keep running.
echo ==================================================
echo.
pause
endlocal
exit /b 0

:diagnose
echo.
echo Startup failed. Current service status:
docker compose ps
echo.
echo Recent service logs:
docker compose logs --tail=80 db api worker web
echo.
echo Suggested fixes:
echo 1. Start Docker Desktop and select Linux containers.
echo 2. Check whether ports 5173 and 8000 are already in use.
echo 3. Run: docker compose build --no-cache
echo 4. Run: docker compose down, then start this file again. Do not use down -v.
goto failed

:failed
echo.
echo ==================================================
echo  Startup did not complete. The window will stay open.
echo  Read the error above, fix it, then run this BAT again.
echo ==================================================
echo.
pause
endlocal
exit /b 1

:start_migration_watcher
if not exist "runtime" mkdir "runtime" >nul 2>nul
if exist "runtime\.migration-watcher.lock" exit /b 0
del /q "runtime\.storage-migration.stop" >nul 2>nul
start "110 storage migration watcher" /min "%ComSpec%" /d /c call "%~f0" --migration-watcher
exit /b 0

:migration_watcher
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"
set "COMPOSE_PROJECT_NAME=documents_workspace"
rem Compose must read the freshly updated .env during a migration. The
rem startup process may still carry the previous path values in its environment.
set "STORAGE_HOST_PATH="
set "POSTGRES_HOST_PATH="
if not exist "runtime" mkdir "runtime" >nul 2>nul
if exist "runtime\.migration-watcher.lock" exit /b 0
mkdir "runtime\.migration-watcher.lock" >nul 2>nul
if errorlevel 1 exit /b 0
>"runtime\.migration-watcher.pid" echo watcher

:migration_watch_loop
if exist "runtime\.storage-migration.stop" (
  del /q "runtime\.storage-migration.stop" >nul 2>nul
  del /q "runtime\.migration-watcher.pid" >nul 2>nul
  rmdir "runtime\.migration-watcher.lock" >nul 2>nul
  exit /b 0
)
if exist "runtime\.storage-migration.request" call :process_storage_migration
timeout /t 2 /nobreak >nul
goto migration_watch_loop

:process_storage_migration
set "NEW_DB_PATH="
set "NEW_STORAGE_PATH="
set "REQUEST_ID="
set /p REQUEST_ID=<"runtime\.storage-migration.request"
set /p NEW_DB_PATH=<"runtime\.storage-migration.db"
set /p NEW_STORAGE_PATH=<"runtime\.storage-migration.storage"
if not defined REQUEST_ID goto migration_bad_request
if not defined NEW_DB_PATH goto migration_bad_request
if not defined NEW_STORAGE_PATH goto migration_bad_request

set "OLD_DB_PATH=./runtime/postgres"
set "OLD_STORAGE_PATH=./runtime/storage"
if exist ".env" for /f "usebackq tokens=1,* delims== eol=#" %%A in (".env") do (
  if "%%A"=="POSTGRES_HOST_PATH" set "OLD_DB_PATH=%%B"
  if "%%A"=="STORAGE_HOST_PATH" set "OLD_STORAGE_PATH=%%B"
)
set "OLD_DB_PATH=%OLD_DB_PATH:"=%"
set "OLD_STORAGE_PATH=%OLD_STORAGE_PATH:"=%"

>"runtime\.storage-migration.status" echo {"id":"!REQUEST_ID!","status":"processing","database_host_path":"!NEW_DB_PATH!","storage_host_path":"!NEW_STORAGE_PATH!","message":"processing"}

set "OLD_DB_SRC=!OLD_DB_PATH:/=\!"
set "OLD_STORAGE_SRC=!OLD_STORAGE_PATH:/=\!"
set "NEW_DB_SRC=!NEW_DB_PATH:/=\!"
set "NEW_STORAGE_SRC=!NEW_STORAGE_PATH:/=\!"
if "!OLD_DB_SRC:~0,2!"==".\" set "OLD_DB_SRC=%~dp0!OLD_DB_SRC:~2!"
if "!OLD_STORAGE_SRC:~0,2!"==".\" set "OLD_STORAGE_SRC=%~dp0!OLD_STORAGE_SRC:~2!"
if "!NEW_DB_SRC:~0,2!"==".\" set "NEW_DB_SRC=%~dp0!NEW_DB_SRC:~2!"
if "!NEW_STORAGE_SRC:~0,2!"==".\" set "NEW_STORAGE_SRC=%~dp0!NEW_STORAGE_SRC:~2!"

docker compose stop api worker web db >"runtime\.storage-migration.log" 2>&1
if errorlevel 1 goto migration_failed
if /i not "!OLD_DB_SRC!"=="!NEW_DB_SRC!" if exist "!NEW_DB_SRC!\*" goto migration_target_exists
if /i not "!OLD_STORAGE_SRC!"=="!NEW_STORAGE_SRC!" if exist "!NEW_STORAGE_SRC!\*" goto migration_target_exists
if not exist "!NEW_DB_SRC!" mkdir "!NEW_DB_SRC!" >nul 2>>"runtime\.storage-migration.log"
if not exist "!NEW_STORAGE_SRC!" mkdir "!NEW_STORAGE_SRC!" >nul 2>>"runtime\.storage-migration.log"

if /i not "!OLD_DB_SRC!"=="!NEW_DB_SRC!" (
  robocopy "!OLD_DB_SRC!" "!NEW_DB_SRC!" /E /COPY:DAT /DCOPY:DAT /R:1 /W:1 /NFL /NDL /NJH /NJS >>"runtime\.storage-migration.log"
  if errorlevel 8 goto migration_failed
)
if /i not "!OLD_STORAGE_SRC!"=="!NEW_STORAGE_SRC!" (
  robocopy "!OLD_STORAGE_SRC!" "!NEW_STORAGE_SRC!" /E /COPY:DAT /DCOPY:DAT /R:1 /W:1 /NFL /NDL /NJH /NJS >>"runtime\.storage-migration.log"
  if errorlevel 8 goto migration_failed
)
call :write_storage_env "!NEW_DB_PATH!" "!NEW_STORAGE_PATH!"
if errorlevel 1 goto migration_failed
docker compose up -d >"runtime\.storage-migration.log" 2>&1
if errorlevel 1 (
  call :write_storage_env "!OLD_DB_PATH!" "!OLD_STORAGE_PATH!"
  docker compose up -d >>"runtime\.storage-migration.log" 2>&1
  goto migration_failed
)
if /i not "!OLD_DB_SRC!"=="!NEW_DB_SRC!" rmdir /s /q "!OLD_DB_SRC!" >nul 2>>"runtime\.storage-migration.log"
if /i not "!OLD_STORAGE_SRC!"=="!NEW_STORAGE_SRC!" rmdir /s /q "!OLD_STORAGE_SRC!" >nul 2>>"runtime\.storage-migration.log"
>"runtime\.storage-migration.status" echo {"id":"!REQUEST_ID!","status":"completed","database_host_path":"!NEW_DB_PATH!","storage_host_path":"!NEW_STORAGE_PATH!","message":"completed"}
del /q "runtime\.storage-migration.request" "runtime\.storage-migration.db" "runtime\.storage-migration.storage" >nul 2>nul
exit /b 0

:migration_target_exists
>"runtime\.storage-migration.status" echo {"id":"!REQUEST_ID!","status":"failed","database_host_path":"!NEW_DB_PATH!","storage_host_path":"!NEW_STORAGE_PATH!","message":"target directory must be empty"}
docker compose up -d >>"runtime\.storage-migration.log" 2>&1
del /q "runtime\.storage-migration.request" "runtime\.storage-migration.db" "runtime\.storage-migration.storage" >nul 2>nul
exit /b 0

:migration_bad_request
>"runtime\.storage-migration.status" echo {"status":"failed","message":"migration request is incomplete"}
del /q "runtime\.storage-migration.request" "runtime\.storage-migration.db" "runtime\.storage-migration.storage" >nul 2>nul
exit /b 0

:migration_failed
>"runtime\.storage-migration.status" echo {"id":"!REQUEST_ID!","status":"failed","database_host_path":"!NEW_DB_PATH!","storage_host_path":"!NEW_STORAGE_PATH!","message":"migration failed; original directories were kept"}
docker compose up -d >>"runtime\.storage-migration.log" 2>&1
del /q "runtime\.storage-migration.request" "runtime\.storage-migration.db" "runtime\.storage-migration.storage" >nul 2>nul
exit /b 0

:write_storage_env
set "ENV_TMP=.env.storage-migration.tmp"
set "FOUND_DB="
set "FOUND_STORAGE="
>"!ENV_TMP!" (
  if exist ".env" for /f "usebackq delims=" %%L in (".env") do (
    for /f "tokens=1,* delims==" %%A in ("%%L") do (
      if "%%A"=="POSTGRES_HOST_PATH" (
        echo POSTGRES_HOST_PATH=%~1
        set "FOUND_DB=1"
      ) else if "%%A"=="STORAGE_HOST_PATH" (
        echo STORAGE_HOST_PATH=%~2
        set "FOUND_STORAGE=1"
      ) else echo(%%L
    )
  )
  if not defined FOUND_DB echo POSTGRES_HOST_PATH=%~1
  if not defined FOUND_STORAGE echo STORAGE_HOST_PATH=%~2
)
move /y "!ENV_TMP!" ".env" >nul 2>nul
if errorlevel 1 exit /b 1
exit /b 0
