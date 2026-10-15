@echo off
setlocal

cd /d "D:\Proyectos\Creador de Excels"

echo Iniciando Codex en D:\Proyectos\Creador de Excels...
echo Red habilitada y acceso completo al sistema de archivos.
echo Las acciones sensibles solicitaran confirmacion.

codex --sandbox danger-full-access --ask-for-approval on-request

endlocal
