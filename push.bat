  @echo off
  chcp 65001 >nul
  cd /d H:\xm4

  echo ========================================
  echo   xm4 项目一键同步到 GitHub
  echo ========================================
  echo.

  set /p msg="请输入本次更新说明（直接回车用默认）: "
  if "%msg%"=="" set msg=update xm4 project

  echo.
  echo [1/3] 添加所有变更...
  git add -A

  echo.
  echo [2/3] 提交变更...
  git commit -m "%msg%"

  echo.
  echo [3/3] 推送到 GitHub...
  git push origin main

  echo.
  echo ========================================
  echo   完成！按任意键退出
  echo ========================================
  pause ^>nul