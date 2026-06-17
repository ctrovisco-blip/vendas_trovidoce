"""
Monitoriza a pasta PASTA_FONTE e gera automaticamente o relatório sempre
que um ficheiro .xlsx for criado ou modificado.

Uso:
    python monitorizar.py

Deixa esta janela aberta. Prima Ctrl+C para parar.
Requer: pip install watchdog
"""

import time
import os
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from gerar_relatorio import PASTA_FONTE, gerar
from datetime import datetime

EXTENSOES_VALIDAS = {".xlsx", ".xls"}
COOLDOWN_SEGUNDOS = 5  # evita disparar várias vezes seguidas


class HandlerExcel(FileSystemEventHandler):
    def __init__(self):
        self._ultimo_evento = 0

    def _processar(self, caminho):
        _, ext = os.path.splitext(caminho)
        if ext.lower() not in EXTENSOES_VALIDAS:
            return
        agora = time.time()
        if agora - self._ultimo_evento < COOLDOWN_SEGUNDOS:
            return
        self._ultimo_evento = agora
        print(f"\n[{datetime.now():%H:%M:%S}] Alteração detectada: {os.path.basename(caminho)}")
        try:
            gerar(caminho)
        except Exception as e:
            print(f"  ERRO: {e}")

    def on_created(self, event):
        if not event.is_directory:
            self._processar(event.src_path)

    def on_modified(self, event):
        if not event.is_directory:
            self._processar(event.src_path)


if __name__ == "__main__":
    print(f"A monitorizar: {PASTA_FONTE}")
    print("Prima Ctrl+C para parar.\n")

    observer = Observer()
    observer.schedule(HandlerExcel(), path=PASTA_FONTE, recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
    print("\nMonitorização terminada.")
