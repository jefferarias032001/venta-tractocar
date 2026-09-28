@echo off
cd /d C:\Users\jarias\Desktop\venta-tractocar
python procesar_ventas.py
git add index.html dashboard_ventas.html procesar_ventas.py
git diff --cached --quiet && goto nochanges
git commit -m "Actualizacion de datos %DATE%  %TIME%"
git push origin master
:nochanges
python enviar_correo.py
