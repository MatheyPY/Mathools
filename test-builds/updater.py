"""
  UPDATER — Mathools
  Sistema de atualização automática via GitHub Releases com intermediário .BAT
  
  Fluxo:
  1. Baixa novo EXE como "_update_new.exe"
  2. Renomeia atual → "Mathools.exe.old"
  3. Renomeia "_update_new.exe" → "Mathools.exe"
  4. Cria "_upd_launch.bat" que aguarda processo antigo morrer, limpa _MEI e abre novo
  5. Lança .bat com Popen (desvinculado) e retorna
  6. launcher_gui.py chama page.window_destroy() para fechar naturalmente
  7. .bat detecta PID sumiu, aguarda 5s, deleta .old, abre novo EXE
"""

import os
import sys
import threading
import subprocess
import logging
import time
from dataclasses import dataclass
from typing import Optional, Callable

# Suprimir SSL warnings quando verify=False (necessário para compatibilidade)
import warnings
warnings.filterwarnings("ignore", message="Unverified HTTPS request")
try:
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
except Exception:
    pass

logger = logging.getLogger(__name__)

# Flag para lazy import de requests (só carrega quando necessário)
_HAS_REQUESTS = None

def _check_requests() -> bool:
    """Verifica e cacheia se requests está disponível (lazy import)."""
    global _HAS_REQUESTS
    if _HAS_REQUESTS is not None:
        return _HAS_REQUESTS
    
    try:
        import requests as _req
        _HAS_REQUESTS = True
        return True
    except ImportError:
        _HAS_REQUESTS = False
        logger.warning("[Updater] Módulo 'requests' não disponível.")
        return False

# ─────────────────────────────────────────────────────────────────────────────
#  CONFIGURAÇÃO
# ─────────────────────────────────────────────────────────────────────────────
CURRENT_VERSION = "2.2.8"
VERSION_JSON_URL = "https://raw.githubusercontent.com/MatheyPY/Mathools/main/version.json"
CHECK_TIMEOUT = 5

@dataclass
class UpdateInfo:
    """Informações sobre atualização disponível."""
    has_update: bool
    latest_version: str
    download_url: str
    release_notes: str
    package_type: str = "portable"  # "portable" (legado) | "installer"
    current_version: str = CURRENT_VERSION


# ─────────────────────────────────────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _parse_version(v: str) -> tuple:
    """Converte versão string para tuple para comparação."""
    try:
        return tuple(int(x) for x in str(v).strip().lstrip("v").split("."))
    except Exception:
        return (0, 0, 0)


def _exe_dir() -> str:
    """Retorna diretório do executável (ou script .py)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _current_exe() -> str:
    """Retorna caminho completo do executável atual (ou script .py)."""
    if getattr(sys, "frozen", False):
        return sys.executable
    return os.path.abspath(__file__)


def _exe_basename() -> str:
    """Retorna só o nome do EXE (ex: 'Mathools.exe')."""
    return os.path.basename(_current_exe())


def _get_disk_space(path: str) -> int:
    """Retorna espaço livre em disco (em bytes) ou -1 se falhar."""
    try:
        import shutil
        stat = shutil.disk_usage(path)
        return stat.free
    except Exception:
        return -1


def _write_error_log(message: str):
    """Escreve erro em arquivo de fallback, sem depender de logging."""
    try:
        import datetime
        timestamp = datetime.datetime.now().isoformat()
        error_file = os.path.join(_exe_dir(), "_updater_error.txt")
        with open(error_file, "a", encoding="utf-8") as f:
            f.write(f"[{timestamp}] {message}\n")
        # Também imprime no stderr para aparecer em consoles/logs externos
        print(f"[UPDATER_LOG][{timestamp}] {message}", flush=True)
    except Exception:
        pass  # Silenciosamente ignora erros de escrita de fallback


def _delete_file_with_retry(path: str, attempts: int = 5, delay: float = 1.0) -> bool:
    """Remove arquivo com múltiplas tentativas, removendo atributos de proteção."""
    if not os.path.exists(path):
        return True
    
    for i in range(attempts):
        try:
            # Tira atributo read-only se existir
            try:
                import stat
                os.chmod(path, stat.S_IWRITE | stat.S_IREAD)
            except Exception:
                pass
            
            os.remove(path)
            logger.info(f"[Updater] Arquivo removido: {path}")
            return True
        except Exception as e:
            if i < attempts - 1:
                logger.debug(f"[Updater] Tentativa {i+1}/{attempts} de remover {os.path.basename(path)}: {type(e).__name__}")
                time.sleep(delay)
    
    logger.warning(f"[Updater] Não foi possível remover após {attempts} tentativas: {path}")
    return False


def _create_bat_launcher(exe_path: str, old_exe_path: str, pid: int) -> Optional[str]:
    try:
        exe_dir = os.path.dirname(exe_path)
        exe_name = os.path.basename(exe_path)
        new_exe = os.path.join(exe_dir, "_update_new.exe")
        bat_path = os.path.join(exe_dir, "_upd_launch.bat")
        vbs_path = os.path.join(exe_dir, "_upd_launch.vbs")

        log_path = os.path.join(exe_dir, "_upd_launch.log")

        bat_lines = [
            "@echo off",
            "setlocal EnableDelayedExpansion",
            f"set LOGFILE={log_path}",
            "",
            "call :LOG \"=== Mathools Updater BAT iniciado ===""",
            f"call :LOG \"PID monitorado : {pid}\"",
            f"call :LOG \"EXE atual      : {exe_path}\"",
            f"call :LOG \"EXE novo       : {new_exe}\"",
            f"call :LOG \"EXE .old       : {old_exe_path}\"",
            "",
            "REM 1. Aguarda PID encerrar",
            f"call :LOG \"[1] Aguardando PID {pid} encerrar...\"",
            "set /A WAIT_COUNT=0",
            ":WAIT_PID",
            "set /A WAIT_COUNT+=1",
            "call :LOG \"[1] Tentativa !WAIT_COUNT!\"",
            f"tasklist /FI \"PID eq {pid}\" /FO CSV /NH 2>nul | find /I \"{pid}\" >nul 2>&1",
            "if not errorlevel 1 (",
            "    call :LOG \"[1] PID ainda vivo. Aguardando 1s...\"",
            "    timeout /T 1 /NOBREAK >nul",
            "    if !WAIT_COUNT! GEQ 120 (",
            "        call :LOG \"[1] TIMEOUT 120s. Abortando.\"",
            "        goto :EOF",
            "    )",
            "    goto WAIT_PID",
            ")",
            "call :LOG \"[1] PID encerrado apos !WAIT_COUNT! tentativas.\"",
            "",
            "REM 2. Aguarda handles DLL",
            "call :LOG \"[2] Aguardando 3s para liberar handles DLL...\"",
            "timeout /T 3 /NOBREAK >nul",
            "",
            "REM 3. Remove .old anterior",
            f"if exist \"{old_exe_path}\" (",
            "    call :LOG \"[3] Removendo .old anterior...\"",
            f"    del /F /Q \"{old_exe_path}\" >nul 2>&1",
            f"    if exist \"{old_exe_path}\" (",
            "        call :LOG \"[3] AVISO: nao conseguiu remover .old\"",
            "    ) else (",
            "        call :LOG \"[3] .old removido com sucesso\"",
            "    )",
            ") else (",
            "    call :LOG \"[3] Nenhum .old anterior encontrado\"",
            ")",
            "",
            "REM 4. Verifica _update_new.exe",
            f"if not exist \"{new_exe}\" (",
            "    call :LOG \"[4] ERRO CRITICO: _update_new.exe NAO encontrado! Abortando.\"",
            "    goto :EOF",
            ")",
            f"for %%F in (\"{new_exe}\") do set NEW_SIZE=%%~zF",
            "call :LOG \"[4] _update_new.exe tamanho: !NEW_SIZE! bytes\"",
            "if !NEW_SIZE! LSS 1000000 (",
            "    call :LOG \"[4] AVISO: arquivo parece pequeno (!NEW_SIZE! bytes)\"",
            ")",
            "",
            "REM 5. Move EXE para .old (retry loop)",
            "call :LOG \"[5] Tentando renomear EXE atual para .old...\"",
            "set /A RENAME_COUNT=0",
            ":RENAME_LOOP",
            "set /A RENAME_COUNT+=1",
            "call :LOG \"[5] Tentativa !RENAME_COUNT! de renomear...\"",
            f"move /Y \"{exe_path}\" \"{old_exe_path}\" >nul 2>&1",
            f"if exist \"{exe_path}\" (",
            "    call :LOG \"[5] EXE ainda bloqueado. Aguardando 2s...\"",
            "    timeout /T 2 /NOBREAK >nul",
            "    if !RENAME_COUNT! GEQ 30 (",
            "        call :LOG \"[5] TIMEOUT: nao conseguiu renomear em 30 tentativas. Abortando.\"",
            "        goto :EOF",
            "    )",
            "    goto RENAME_LOOP",
            ")",
            "call :LOG \"[5] EXE renomeado para .old na tentativa !RENAME_COUNT!\"",
            "",
            "REM 6. Move _update_new.exe para posicao oficial",
            "call :LOG \"[6] Movendo _update_new.exe para posicao oficial...\"",
            f"move /Y \"{new_exe}\" \"{exe_path}\" >nul 2>&1",
            f"if not exist \"{exe_path}\" (",
            "    call :LOG \"[6] ERRO CRITICO: move falhou! Tentando restaurar .old...\"",
            f"    move /Y \"{old_exe_path}\" \"{exe_path}\" >nul 2>&1",
            "    goto :EOF",
            ")",
            "call :LOG \"[6] EXE novo no lugar com sucesso.\"",
            "",
            "REM 7. Nao limpar _MEI aqui (evita corrida com bootstrap do PyInstaller)",
            "call :LOG \"[7] Pulando limpeza de _MEI para evitar erro de python312.dll\"",
            "",
            "REM 8. Lanca o novo EXE com retry (evita falha transitória no bootstrap)",
            "call :LOG \"[8] Aguardando 5s final...\"",
            "timeout /T 5 /NOBREAK >nul",
            f"cd /D \"{exe_dir}\"",
            f"set \"TEMP={exe_dir}\"",
            f"set \"TMP={exe_dir}\"",
            "call :LOG \"[8] TEMP/TMP redirecionados para pasta do app.\"",
            "set /A LAUNCH_ATTEMPT=0",
            "set LAUNCH_OK=0",
            ":LAUNCH_RETRY",
            "set /A LAUNCH_ATTEMPT+=1",
            f"call :LOG \"[8] Tentativa !LAUNCH_ATTEMPT! de iniciar {exe_name}\"",
            f"start \"Mathools\" \"{exe_path}\"",
            "timeout /T 4 /NOBREAK >nul",
            f"tasklist /FI \"IMAGENAME eq {exe_name}\" /FO CSV /NH 2>nul | find /I \"{exe_name}\" >nul 2>&1",
            "if not errorlevel 1 (",
            "    call :LOG \"[8] Novo EXE em execucao.\"",
            "    set LAUNCH_OK=1",
            ") else (",
            "    call :LOG \"[8] AVISO: processo nao encontrado apos start.\"",
            "    if !LAUNCH_ATTEMPT! LSS 3 (",
            "        call :LOG \"[8] Repetindo tentativa em 3s...\"",
            "        timeout /T 3 /NOBREAK >nul",
            "        goto LAUNCH_RETRY",
            "    )",
            "    call :LOG \"[8] ERRO: falha apos 3 tentativas de inicializacao.\"",
            ")",
            "",
            "REM 9. Valida estabilidade do novo EXE e faz rollback se necessário",
            "if \"!LAUNCH_OK!\"==\"1\" (",
            "    call :LOG \"[9] Validando estabilidade por 12s...\"",
            "    timeout /T 12 /NOBREAK >nul",
            f"    tasklist /FI \"IMAGENAME eq {exe_name}\" /FO CSV /NH 2>nul | find /I \"{exe_name}\" >nul 2>&1",
            "    if not errorlevel 1 (",
            "        set MEI_OK=0",
            f"        for /D %%D in (\"{exe_dir}\\_MEI*\") do (",
            "            if exist \"%%D\\python312.dll\" set MEI_OK=1",
            "        )",
            "        if \"!MEI_OK!\"==\"1\" (",
            "            call :LOG \"[9] Novo EXE permaneceu em execucao e python312.dll encontrado no _MEI (estavel).\"",
            "        ) else (",
            "            call :LOG \"[9] ERRO: processo vivo, mas python312.dll ausente no _MEI. Tentando rollback...\"",
            f"            if exist \"{old_exe_path}\" (",
            f"                del /F /Q \"{exe_path}\" >nul 2>&1",
            f"                move /Y \"{old_exe_path}\" \"{exe_path}\" >nul 2>&1",
            f"                if exist \"{exe_path}\" (",
            "                    call :LOG \"[9] Rollback concluido. Relancando versao antiga...\"",
            f"                    start \"Mathools\" \"{exe_path}\"",
            "                    goto :FINAL_CLEAN",
            "                )",
            "            )",
            "            call :LOG \"[9] Rollback indisponivel/falhou.\"",
            "            goto :FINAL_CLEAN",
            "        )",
            "    ) else (",
            "        call :LOG \"[9] ERRO: novo EXE caiu apos iniciar. Tentando rollback...\"",
            f"        if exist \"{old_exe_path}\" (",
            f"            del /F /Q \"{exe_path}\" >nul 2>&1",
            f"            move /Y \"{old_exe_path}\" \"{exe_path}\" >nul 2>&1",
            f"            if exist \"{exe_path}\" (",
            "                call :LOG \"[9] Rollback concluido. Relancando versao antiga...\"",
            f"                start \"Mathools\" \"{exe_path}\"",
            "                goto :FINAL_CLEAN",
            "            )",
            "        )",
            "        call :LOG \"[9] Rollback indisponivel/falhou.\"",
            "        goto :FINAL_CLEAN",
            "    )",
            ") else (",
            "    call :LOG \"[9] Novo EXE nao iniciou; mantendo .old para recuperacao manual.\"",
            "    goto :FINAL_CLEAN",
            ")",
            "",
            "REM 10. Remove versão antiga (.old)",
            f"if exist \"{old_exe_path}\" (",
            "    call :LOG \"[10] Removendo EXE antigo (.old)...\"",
            f"    del /F /Q \"{old_exe_path}\" >nul 2>&1",
            f"    if exist \"{old_exe_path}\" (",
            "        call :LOG \"[10] AVISO: nao conseguiu remover .old (tentara no proximo inicio)\"",
            "    ) else (",
            "        call :LOG \"[10] EXE antigo removido com sucesso\"",
            "    )",
            ")",
            "",
            ":FINAL_CLEAN",
            "REM 11. Limpeza final",
            "call :LOG \"[11] Removendo scripts de atualizacao...\"",
            "call :LOG \"=== BAT concluido com sucesso ===\"",
            f"del /F /Q \"{vbs_path}\" >nul 2>&1",
            "del /F /Q \"%~f0\" >nul 2>&1",
            "goto :EOF",
            "",
            ":LOG",
            "echo [%DATE% %TIME%] %~1 >> \"%LOGFILE%\"",
            "goto :EOF",
        ]
        bat_content = "\r\n".join(bat_lines) + "\r\n"

        vbs_lines = [
            "On Error Resume Next",
            "Dim objShell, strBatPath",
            "Set objShell = CreateObject(\"WScript.Shell\")",
            f"strBatPath = \"{bat_path}\"",
            "objShell.Run chr(34) & strBatPath & chr(34), 0, True",
            "Set objShell = Nothing",
        ]
        vbs_content = "\n".join(vbs_lines) + "\n"

        with open(bat_path, "w", encoding="cp1252", errors="replace") as f:
            f.write(bat_content)
        bat_size = os.path.getsize(bat_path)
        logger.debug(f"[Updater][launcher] BAT escrito: {bat_path} ({bat_size} bytes)")
        logger.debug(f"[Updater][launcher] Conteudo do BAT:\n{bat_content}")
        _write_error_log(f"[launcher] BAT={bat_path} ({bat_size} bytes)")

        with open(vbs_path, "w", encoding="cp1252", errors="replace") as f:
            f.write(vbs_content)
        vbs_size = os.path.getsize(vbs_path)
        logger.debug(f"[Updater][launcher] VBS escrito: {vbs_path} ({vbs_size} bytes)")
        _write_error_log(f"[launcher] VBS={vbs_path} ({vbs_size} bytes)")

        # Verifica se os arquivos existem e têm conteúdo
        if bat_size == 0:
            logger.error(f"[Updater][launcher] BAT foi criado mas está VAZIO!")
            _write_error_log("[launcher] ERRO: BAT vazio")
            return None
        if vbs_size == 0:
            logger.error(f"[Updater][launcher] VBS foi criado mas está VAZIO!")
            _write_error_log("[launcher] ERRO: VBS vazio")
            return None

        logger.info(f"[Updater][launcher] Scripts criados com sucesso: BAT={bat_size}b | VBS={vbs_size}b")
        return vbs_path

    except Exception as e:
        import traceback
        msg = f"Erro ao criar launcher: {type(e).__name__}: {e}\n{traceback.format_exc()}"
        logger.error(f"[Updater] {msg}")
        _write_error_log(f"[launcher] ERRO: {msg}")
        return None


def _create_installer_launcher(installer_path: str, exe_path: str, pid: int) -> Optional[str]:
    """Cria launcher para executar instalador silencioso após o app encerrar."""
    try:
        exe_dir = os.path.dirname(exe_path)
        exe_name = os.path.basename(exe_path)
        bat_path = os.path.join(exe_dir, "_upd_launch.bat")
        vbs_path = os.path.join(exe_dir, "_upd_launch.vbs")
        log_path = os.path.join(exe_dir, "_upd_launch.log")

        bat_lines = [
            "@echo off",
            "setlocal EnableDelayedExpansion",
            f"set LOGFILE={log_path}",
            f"set INSTALL_LOG={os.path.join(exe_dir, '_upd_install.log')}",
            "",
            "call :LOG \"=== Mathools Installer Updater iniciado ===\"",
            f"call :LOG \"PID monitorado : {pid}\"",
            f"call :LOG \"Instalador    : {installer_path}\"",
            f"call :LOG \"Destino       : {exe_dir}\"",
            "",
            "REM 1. Aguarda PID encerrar",
            "set /A WAIT_COUNT=0",
            ":WAIT_PID",
            "set /A WAIT_COUNT+=1",
            f"tasklist /FI \"PID eq {pid}\" /FO CSV /NH 2>nul | find /I \"{pid}\" >nul 2>&1",
            "if not errorlevel 1 (",
            "    timeout /T 1 /NOBREAK >nul",
            "    if !WAIT_COUNT! GEQ 15 (",
            "        call :LOG \"[1] TIMEOUT aguardando fechamento do app. Forcando encerramento do PID monitorado...\"",
            f"        taskkill /PID {pid} /F >nul 2>&1",
            "        timeout /T 2 /NOBREAK >nul",
            f"        tasklist /FI \"PID eq {pid}\" /FO CSV /NH 2>nul | find /I \"{pid}\" >nul 2>&1",
            "        if not errorlevel 1 (",
            "            call :LOG \"[1] ERRO: nao foi possivel encerrar o processo antigo.\"",
            "            goto :FINAL",
            "        )",
            "        call :LOG \"[1] Processo antigo encerrado via taskkill.\"",
            "        goto :PID_DONE",
            "    )",
            "    goto WAIT_PID",
            ")",
            ":PID_DONE",
            "call :LOG \"[1] Processo antigo encerrado.\"",
            "",
            "REM 2. Executa instalador silencioso no mesmo diretório",
            f"if not exist \"{installer_path}\" (",
            "    call :LOG \"[2] ERRO: instalador nao encontrado.\"",
            "    goto :FINAL",
            ")",
            "call :LOG \"[2] Executando instalador silencioso...\"",
            f"start /wait \"\" \"{installer_path}\" /VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP- /DIR=\"{exe_dir}\" /LOG=\"{os.path.join(exe_dir, '_upd_install.log')}\"",
            "set INSTALL_RC=!ERRORLEVEL!",
            "call :LOG \"[2] Instalador finalizado. RC=!INSTALL_RC!\"",
            "if not \"!INSTALL_RC!\"==\"0\" (",
            "    call :LOG \"[2] AVISO: instalador retornou codigo diferente de 0.\"",
            ")",
            "",
            "REM 3. Limpa instalador temporário",
            f"del /F /Q \"{installer_path}\" >nul 2>&1",
            "",
            "REM 4. Relança app (com retry e fallback)",
            "set \"TARGET_EXE=\"",
            f"if exist \"{exe_path}\" set \"TARGET_EXE={exe_path}\"",
            "if not defined TARGET_EXE if exist \"%LOCALAPPDATA%\\Mathools\\Mathools.exe\" set \"TARGET_EXE=%LOCALAPPDATA%\\Mathools\\Mathools.exe\"",
            "if not defined TARGET_EXE (",
            "    call :LOG \"[4] ERRO: app nao encontrado apos instalacao.\"",
            "    goto :FINAL",
            ")",
            "call :LOG \"[4] Tentando relancar: %TARGET_EXE%\"",
            "set /A LAUNCH_COUNT=0",
            ":LAUNCH_RETRY",
            "set /A LAUNCH_COUNT+=1",
            "start \"Mathools\" \"%TARGET_EXE%\"",
            "timeout /T 3 /NOBREAK >nul",
            f"tasklist /FI \"IMAGENAME eq {exe_name}\" /FO CSV /NH 2>nul | find /I \"{exe_name}\" >nul 2>&1",
            "if not errorlevel 1 (",
            "    call :LOG \"[4] App relancado com sucesso (tentativa !LAUNCH_COUNT!).\"",
            ") else (",
            "    call :LOG \"[4] AVISO: app nao subiu (tentativa !LAUNCH_COUNT!).\"",
            "    if !LAUNCH_COUNT! LSS 3 (",
            "        timeout /T 2 /NOBREAK >nul",
            "        goto :LAUNCH_RETRY",
            "    )",
            "    call :LOG \"[4] ERRO: falha ao relancar apos 3 tentativas.\"",
            "    )",
            "",
            "",
            ":FINAL",
            f"del /F /Q \"{vbs_path}\" >nul 2>&1",
            "del /F /Q \"%~f0\" >nul 2>&1",
            "goto :EOF",
            "",
            ":LOG",
            "echo [%DATE% %TIME%] %~1 >> \"%LOGFILE%\"",
            "goto :EOF",
        ]
        bat_content = "\r\n".join(bat_lines) + "\r\n"

        vbs_lines = [
            "On Error Resume Next",
            "Dim objShell, strBatPath",
            "Set objShell = CreateObject(\"WScript.Shell\")",
            f"strBatPath = \"{bat_path}\"",
            "objShell.Run chr(34) & strBatPath & chr(34), 0, True",
            "Set objShell = Nothing",
        ]
        vbs_content = "\n".join(vbs_lines) + "\n"

        with open(bat_path, "w", encoding="cp1252", errors="replace") as f:
            f.write(bat_content)
        with open(vbs_path, "w", encoding="cp1252", errors="replace") as f:
            f.write(vbs_content)
        return vbs_path
    except Exception as e:
        _write_error_log(f"[installer_launcher] ERRO: {type(e).__name__}: {e}")
        return None


def _create_script_post_install_launcher(installer_pid: int, target_exe: str, work_dir: str) -> Optional[str]:
    """Cria launcher para aguardar o instalador silencioso e relançar o app instalado."""
    try:
        exe_name = os.path.basename(target_exe)
        bat_path = os.path.join(work_dir, "_upd_after_install.bat")
        vbs_path = os.path.join(work_dir, "_upd_after_install.vbs")
        log_path = os.path.join(work_dir, "_upd_after_install.log")

        bat_lines = [
            "@echo off",
            "setlocal EnableDelayedExpansion",
            f"set LOGFILE={log_path}",
            "",
            "call :LOG \"=== Mathools Script Post-Install iniciado ===\"",
            f"call :LOG \"PID instalador : {installer_pid}\"",
            f"call :LOG \"Target         : {target_exe}\"",
            "",
            "set /A WAIT_COUNT=0",
            ":WAIT_INSTALLER",
            "set /A WAIT_COUNT+=1",
            f"tasklist /FI \"PID eq {installer_pid}\" /FO CSV /NH 2>nul | find /I \"{installer_pid}\" >nul 2>&1",
            "if not errorlevel 1 (",
            "    if !WAIT_COUNT! LEQ 120 (",
            "        timeout /T 1 /NOBREAK >nul",
            "        goto WAIT_INSTALLER",
            "    )",
            "    call :LOG \"[1] TIMEOUT aguardando instalador encerrar.\"",
            "    goto :FINAL",
            ")",
            "call :LOG \"[1] Instalador encerrado.\"",
            "timeout /T 2 /NOBREAK >nul",
            f"if not exist \"{target_exe}\" (",
            "    call :LOG \"[2] ERRO: executavel instalado nao encontrado.\"",
            "    goto :FINAL",
            ")",
            "call :LOG \"[2] Tentando relancar app instalado...\"",
            "set /A LAUNCH_COUNT=0",
            ":LAUNCH_RETRY",
            "set /A LAUNCH_COUNT+=1",
            f"start \"Mathools\" \"{target_exe}\"",
            "timeout /T 3 /NOBREAK >nul",
            f"tasklist /FI \"IMAGENAME eq {exe_name}\" /FO CSV /NH 2>nul | find /I \"{exe_name}\" >nul 2>&1",
            "if not errorlevel 1 (",
            "    call :LOG \"[2] App relancado com sucesso (tentativa !LAUNCH_COUNT!).\"",
            "    goto :FINAL",
            ")",
            "call :LOG \"[2] AVISO: app nao subiu (tentativa !LAUNCH_COUNT!).\"",
            "if !LAUNCH_COUNT! LSS 3 (",
            "    timeout /T 2 /NOBREAK >nul",
            "    goto :LAUNCH_RETRY",
            ")",
            "call :LOG \"[2] ERRO: falha ao relancar apos 3 tentativas.\"",
            "",
            ":FINAL",
            f"del /F /Q \"{vbs_path}\" >nul 2>&1",
            "del /F /Q \"%~f0\" >nul 2>&1",
            "goto :EOF",
            "",
            ":LOG",
            "echo [%DATE% %TIME%] %~1 >> \"%LOGFILE%\"",
            "goto :EOF",
        ]
        bat_content = "\r\n".join(bat_lines) + "\r\n"

        vbs_lines = [
            "On Error Resume Next",
            "Dim objShell, strBatPath",
            "Set objShell = CreateObject(\"WScript.Shell\")",
            f"strBatPath = \"{bat_path}\"",
            "objShell.Run chr(34) & strBatPath & chr(34), 0, False",
            "Set objShell = Nothing",
        ]
        vbs_content = "\n".join(vbs_lines) + "\n"

        with open(bat_path, "w", encoding="cp1252", errors="replace") as f:
            f.write(bat_content)
        with open(vbs_path, "w", encoding="cp1252", errors="replace") as f:
            f.write(vbs_content)
        return vbs_path
    except Exception as e:
        _write_error_log(f"[script_post_install_launcher] ERRO: {type(e).__name__}: {e}")
        return None


# ─────────────────────────────────────────────────────────────────────────────
#  CLASSES PRINCIPAIS
# ─────────────────────────────────────────────────────────────────────────────

class UpdateChecker:
    """Verificador de atualizações e gerenciador de instalação."""
    
    def __init__(self):
        """
        Inicializa e executa limpeza silenciosa (faxineiro).
        Se o novo processo iniciou após atualização, remove arquivos deixados.
        Não faz logging aqui para não interferir com inicialização.
        """
        if not getattr(sys, "frozen", False):
            return
        
        # Faxineiro silencioso — sem logging aqui
        try:
            exe_dir = _exe_dir()
            exe_basename = _exe_basename()
            old_exe_path = os.path.join(exe_dir, f"{exe_basename}.old")
            bat_path = os.path.join(exe_dir, "_upd_launch.bat")
            vbs_path = os.path.join(exe_dir, "_upd_launch.vbs")

            # Tenta deletar .old deixado pelo launcher anterior
            if os.path.exists(old_exe_path):
                _delete_file_with_retry(old_exe_path, attempts=5, delay=1.0)

            # Tenta deletar .bat e .vbs deixados pelo launcher
            if os.path.exists(bat_path):
                _delete_file_with_retry(bat_path, attempts=3, delay=0.5)
            if os.path.exists(vbs_path):
                _delete_file_with_retry(vbs_path, attempts=3, delay=0.5)
        except Exception:
            # Ignora erros silenciosamente na inicialização
            pass

    def check(self) -> Optional[UpdateInfo]:
        """
        Verifica se há atualização disponível.
        
        Retorna:
            UpdateInfo com has_update=True se há nova versão
            None se houve erro na verificação
        """
        if not _check_requests():
            logger.warning("[Updater] 'requests' não disponível — verificação ignorada.")
            return None

        if "SEU_USUARIO" in VERSION_JSON_URL or not VERSION_JSON_URL.startswith("http"):
            logger.warning("[Updater] VERSION_JSON_URL não configurada — verificação ignorada.")
            return None

        logger.debug(f"[Updater][check] Consultando URL: {VERSION_JSON_URL}")
        logger.debug(f"[Updater][check] Versão atual: {CURRENT_VERSION} | timeout: {CHECK_TIMEOUT}s")
        try:
            import requests
            cache_buster = int(time.time())
            sep = "&" if "?" in VERSION_JSON_URL else "?"
            check_url = f"{VERSION_JSON_URL}{sep}t={cache_buster}"
            resp = requests.get(
                check_url,
                timeout=CHECK_TIMEOUT,
                verify=False,
                headers={"Cache-Control": "no-cache", "Pragma": "no-cache"},
            )
            logger.debug(f"[Updater][check] HTTP status: {resp.status_code}")
            logger.debug(f"[Updater][check] Content-Type: {resp.headers.get('Content-Type')}")
            resp.raise_for_status()
            raw_text = resp.text
            logger.debug(f"[Updater][check] Resposta bruta: {raw_text[:300]}")
            data = resp.json()
            logger.debug(f"[Updater][check] JSON parseado: {data}")
        except Exception as e:
            import traceback
            msg = f"Não foi possível verificar atualizações: {type(e).__name__}: {e}\n{traceback.format_exc()}"
            logger.warning(f"[Updater] {msg}")
            _write_error_log(f"[check] {msg}")
            return None

        latest = str(data.get("version", "0.0.0"))
        dl_url = str(data.get("download_url", ""))
        notes  = str(data.get("release_notes", ""))
        package_type = str(data.get("package_type", "portable")).strip().lower() or "portable"
        
        has_update = _parse_version(latest) > _parse_version(CURRENT_VERSION)
        logger.debug(f"[Updater][check] latest={latest} | dl_url={dl_url} | has_update={has_update}")
        logger.debug(f"[Updater][check] parsed_latest={_parse_version(latest)} | parsed_current={_parse_version(CURRENT_VERSION)}")
        _write_error_log(f"[check] OK | current={CURRENT_VERSION} | latest={latest} | has_update={has_update} | url={dl_url}")

        return UpdateInfo(
            has_update=has_update,
            latest_version=latest,
            download_url=dl_url,
            release_notes=notes,
            package_type=package_type,
            current_version=CURRENT_VERSION,
        )

    def check_async(self, callback: Callable[[Optional[UpdateInfo]], None]) -> threading.Thread:
        """
        Verifica atualização em thread separada (não bloqueia UI).
        
        Args:
            callback: função callable(UpdateInfo | None) chamada ao terminar
        
        Returns:
            threading.Thread da verificação
        """
        def _run():
            info = self.check()
            try:
                callback(info)
            except Exception as e:
                logger.warning(f"[Updater] Erro no callback: {e}")

        t = threading.Thread(target=_run, daemon=True, name="UpdateChecker")
        t.start()
        return t

    def download_and_apply(
        self,
        info: UpdateInfo,
        progress_callback: Optional[Callable[[int, int], None]] = None,
    ) -> bool:
        """
        Baixa e aplica atualização.
        
        Fluxo:
        1. Baixa novo EXE
        2. Substitui arquivo atual
        3. Cria .bat intermediário
        4. Lança .bat desvinculado
        5. Retorna True (sem fechar processo — launcher_gui.py cuida disso)
        
        Args:
            info: UpdateInfo com URL de download
            progress_callback: callable(downloaded, total) para tracking de progresso
        
        Returns:
            True se download e preparação bem-sucedidos
            False se erro em qualquer etapa
        """
        # Log CRÍTICO no início para diagnosticar
        msg_inicio = f"download_and_apply() CHAMADO | has_update={info.has_update}"
        _write_error_log(msg_inicio)
        logger.critical(f"[Updater] {msg_inicio}")
        
        # Try/except abrangente para pegar QUALQUER erro não previsto
        try:
            return self._download_and_apply_impl(info, progress_callback)
        except Exception as e:
            import traceback
            error_msg = f"ERRO CRÍTICO (não capturado): {type(e).__name__}: {e}\n{traceback.format_exc()}"
            logger.critical(f"[Updater] {error_msg}")
            _write_error_log(error_msg)
            return False

    def _download_and_apply_impl(
        self,
        info: UpdateInfo,
        progress_callback: Optional[Callable[[int, int], None]] = None,
    ) -> bool:
        """Implementação interna de download_and_apply com try/excepts específicos."""
        
        # Log CRÍTICO no início para garantir que a função está sendo chamada
        msg_inicio = f"_download_and_apply_impl INICIADO | info.has_update={info.has_update} | info.download_url={info.download_url}"
        _write_error_log(msg_inicio)
        logger.critical(f"[Updater] {msg_inicio}")
        
        # ── SNAPSHOT DO AMBIENTE ─────────────────────────────────────────────
        import platform, datetime
        env_info = (
            f"OS: {platform.system()} {platform.release()} | "
            f"Python: {sys.version.split()[0]} | "
            f"frozen: {getattr(sys, 'frozen', False)} | "
            f"exe: {_current_exe()} | "
            f"cwd: {os.getcwd()} | "
            f"pid: {os.getpid()} | ppid: {os.getppid()} | "
            f"hora: {datetime.datetime.now().isoformat()}"
        )
        logger.debug(f"[Updater][env] {env_info}")
        _write_error_log(f"[env] {env_info}")

        # ── DISCO ─────────────────────────────────────────────────────────────
        disk_free = _get_disk_space(_exe_dir())
        logger.debug(f"[Updater][env] Espaço livre em disco: {disk_free:,} bytes ({disk_free // (1024*1024)} MB)")
        _write_error_log(f"[env] Disco livre: {disk_free:,} bytes")

        if not _check_requests() or not info.download_url:
            msg = f"Configuração inválida | requests={_check_requests()} | download_url={repr(info.download_url)}"
            _write_error_log(msg)
            logger.error(f"[Updater] {msg}")
            return False

        current_exe = _current_exe()
        exe_dir = _exe_dir()
        exe_basename = _exe_basename()
        is_installer_package = str(getattr(info, "package_type", "portable")).lower() == "installer"
        new_exe = os.path.join(exe_dir, "_update_installer.exe" if is_installer_package else "_update_new.exe")
        old_exe = os.path.join(exe_dir, f"{exe_basename}.old")

        logger.info(f"[Updater] ≡≡≡ INICIANDO ATUALIZAÇÃO ≡≡≡")
        logger.info(f"[Updater] Versão: {CURRENT_VERSION} → {info.latest_version}")
        logger.debug(f"[Updater][paths] current_exe={current_exe}")
        logger.debug(f"[Updater][paths] exe_dir={exe_dir}")
        logger.debug(f"[Updater][paths] new_exe={new_exe}")
        logger.debug(f"[Updater][paths] old_exe={old_exe}")
        logger.debug(f"[Updater][paths] new_exe já existe {os.path.exists(new_exe)}")
        logger.debug(f"[Updater][paths] old_exe já existe {os.path.exists(old_exe)}")
        _write_error_log(f"[paths] current={current_exe} | new={new_exe} | old={old_exe}")

        # ────────────────────────────────────────────────────────────────────
        # 1. DOWNLOAD
        # ────────────────────────────────────────────────────────────────────
        try:
            import requests
            from requests.adapters import HTTPAdapter
            from urllib3.util.retry import Retry
            
            logger.info(f"[Updater] [1/5] Baixando nova versão...")
            logger.debug(f"[Updater] URL: {info.download_url}")
            logger.debug(f"[Updater] Diretório de destino: {exe_dir}")
            logger.debug(f"[Updater] Arquivo temporário: {new_exe}")
            
            # Verifica permissões de escrita na pasta
            if not os.access(exe_dir, os.W_OK):
                logger.error(f"[Updater] Sem permissão de escrita em: {exe_dir}")
                return False
            
            # Cria session com retry automático
            session = requests.Session()
            retry_strategy = Retry(
                total=3,
                backoff_factor=1,
                status_forcelist=[429, 500, 502, 503, 504],
            )
            adapter = HTTPAdapter(max_retries=retry_strategy)
            session.mount("http://", adapter)
            session.mount("https://", adapter)
            
            # Cabeçalhos robustos simulando um navegador real (Disfarce contra Firewalls)
            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
                "Accept": "*/*",
            }
            
            logger.debug(f"[Updater][download] Iniciando GET para: {info.download_url}")
            logger.debug(f"[Updater][download] Headers enviados: {headers}")
            _write_error_log(f"[download] GET {info.download_url}")
            with session.get(
                info.download_url,
                stream=True,
                timeout=120,
                allow_redirects=True,
                headers=headers,
                verify=False
            ) as resp:
                resp.raise_for_status()
                logger.debug(f"[Updater][download] HTTP {resp.status_code} | URL final (após redirect): {resp.url}")
                logger.debug(f"[Updater][download] Content-Length: {resp.headers.get('content-length')} | Content-Type: {resp.headers.get('Content-Type')} | Content-Encoding: {resp.headers.get('Content-Encoding')}")
                logger.debug(f"[Updater][download] Headers completos da resposta: {dict(resp.headers)}")
                _write_error_log(f"[download] HTTP {resp.status_code} | url_final={resp.url} | content-length={resp.headers.get('content-length')} | content-type={resp.headers.get('Content-Type')}")

                # >>> VERIFICAÇÃO DE FIREWALL <<<
                # Se a resposta vier como HTML em vez de arquivo executável, bloqueia.
                content_type = resp.headers.get('Content-Type', '').lower()
                logger.debug(f"[Updater][download] Content-Type verificado: '{content_type}'")
                if 'text/html' in content_type:
                    # Pega os primeiros 500 bytes para diagnóstico
                    preview = resp.content[:500] if hasattr(resp, 'content') else b'(nao disponivel)'
                    msg = f"Download interceptado (content-type='{content_type}'). Preview: {preview}"
                    logger.error(f"[Updater] {msg}")
                    _write_error_log(msg)
                    return False

                total = int(resp.headers.get("content-length", 0))
                downloaded = 0
                
                # Cria arquivo com modo binário exclusivo (falha se já existir)
                try:
                    # Remove arquivo temporário se existir de atualização anterior
                    if os.path.exists(new_exe):
                        try:
                            os.remove(new_exe)
                            logger.debug(f"[Updater] Arquivo temporário anterior removido")
                        except Exception as e:
                            logger.warning(f"[Updater] Falha ao remover arquivo temporário: {e}")
                    
                    logger.debug(f"[Updater] Abrindo arquivo para escrita: {new_exe}")
                    _last_pct_logged = [-1]  # lista para mutabilidade no closure
                    with open(new_exe, "wb") as f:
                        for chunk in resp.iter_content(chunk_size=65536):
                            if chunk:
                                f.write(chunk)
                                downloaded += len(chunk)
                                if progress_callback and total > 0:
                                    try:
                                        progress_callback(downloaded, total)
                                    except Exception as e:
                                        logger.debug(f"[Updater][download] Erro no progress_callback: {type(e).__name__}: {e}")
                                # Loga a cada 10% (ou a cada chunk se sem content-length)
                                if total > 0:
                                    pct = int(downloaded * 100 / total)
                                    if pct // 10 > _last_pct_logged[0] // 10:
                                        logger.debug(f"[Updater][download] Progresso: {downloaded:,}/{total:,} bytes ({pct}%)")
                                        _write_error_log(f"[download] {pct}% — {downloaded:,}/{total:,} bytes")
                                        _last_pct_logged[0] = pct
                                else:
                                    # Sem content-length: loga a cada ~5 MB
                                    if downloaded % (5 * 1024 * 1024) < 65536:
                                        logger.debug(f"[Updater][download] Recebido até agora: {downloaded:,} bytes (sem content-length)")

                        # Garante que o buffer seja descarregado fisicamente no disco
                        f.flush()
                        os.fsync(f.fileno())
                        logger.debug(f"[Updater][download] flush+fsync OK — {downloaded:,} bytes escritos no buffer")

                    # Verifica tamanho REAL do arquivo no disco após fechar
                    file_size = os.path.getsize(new_exe)
                    logger.info(f"[Updater][download] CONCLUÍDO: {new_exe}")
                    logger.info(f"[Updater][download] Bytes recebidos (stream): {downloaded:,} | Tamanho no disco: {file_size:,} | content-length esperado: {total if total > 0 else '(não informado)'}")
                    _write_error_log(f"[download] CONCLUÍDO | stream={downloaded:,} | disco={file_size:,} | content-length={total}")

                    # Valida: se content-length foi informado, compara com o baixado
                    if total > 0 and file_size < total:
                        msg = f"Download INCOMPLETO: content-length={total:,} | disco={file_size:,} | faltando={total-file_size:,} bytes"
                        logger.error(f"[Updater] {msg}")
                        _write_error_log(msg)
                        try:
                            os.remove(new_exe)
                        except Exception:
                            pass
                        return False

                    # Se não teve content-length: falha em arquivo vazio, avisa se menor que 1 MB
                    if file_size == 0:
                        msg = f"Download resultou em arquivo vazio (0 bytes)"
                        logger.error(f"[Updater] {msg}")
                        _write_error_log(msg)
                        try:
                            os.remove(new_exe)
                        except Exception:
                            pass
                        return False
                    if total == 0 and file_size < 1_000_000:
                        logger.warning(f"[Updater] Arquivo parece pequeno ({file_size} bytes) — continuando mesmo assim")
                
                except IOError as e:
                    import traceback as _tb
                    logger.error(f"[Updater] ERRO ao escrever arquivo: {e}\n{_tb.format_exc()}")
                    logger.debug(f"[Updater] Caminho: {new_exe}")
                    logger.debug(f"[Updater] Espaço em disco disponível: {_get_disk_space(exe_dir)} bytes")
                    _write_error_log(f"IOError ao gravar arquivo: {e}")
                    return False

        except requests.exceptions.ConnectionError as e:
            msg = f"ERRO de conexão: {e}"
            logger.error(f"[Updater] {msg}")
            logger.error("[Updater] Verifique sua conexão de rede ou firewall/proxy corporativo")
            _write_error_log(msg)
            try:
                if os.path.exists(new_exe):
                    os.remove(new_exe)
            except Exception:
                pass
            return False
        except requests.exceptions.Timeout as e:
            msg = f"TIMEOUT no download (timeout=120s): {e}"
            logger.error(f"[Updater] {msg}")
            _write_error_log(msg)
            try:
                if os.path.exists(new_exe):
                    os.remove(new_exe)
            except Exception:
                pass
            return False
        except Exception as e:
            import traceback
            msg = f"ERRO no download: {type(e).__name__}: {e}\nTraceback: {traceback.format_exc()}"
            logger.error(f"[Updater] {msg}")
            logger.debug(f"[Updater] URL tentada: {info.download_url}")
            _write_error_log(msg)
            try:
                if os.path.exists(new_exe):
                    os.remove(new_exe)
            except Exception:
                pass
            return False

        # Modo script .py: como não há bootstrap do PyInstaller para substituir,
        # dispara o instalador baixado e deixa a UI encerrar o processo atual.
        if not getattr(sys, "frozen", False):
            if is_installer_package:
                try:
                    install_dir = os.path.join(os.environ.get("LOCALAPPDATA", exe_dir), "Mathools")
                    launch_args = [
                        new_exe,
                        "/VERYSILENT",
                        "/SUPPRESSMSGBOXES",
                        "/NORESTART",
                        "/SP-",
                        f"/DIR={install_dir}",
                        f"/LOG={os.path.join(install_dir, '_upd_install.log')}",
                    ]
                    logger.info("[Updater] [!] Rodando em script .py — iniciando instalador baixado.")
                    logger.debug(f"[Updater][script] Executando instalador: {launch_args}")
                    _write_error_log(f"[script] Iniciando instalador diretamente: {new_exe}")
                    os.makedirs(install_dir, exist_ok=True)
                    proc = subprocess.Popen(
                        launch_args,
                        cwd=exe_dir,
                        creationflags=0x00000008 | 0x08000000,
                        stdin=subprocess.DEVNULL,
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL,
                        close_fds=True,
                    )
                    _write_error_log(f"[script] Instalador iniciado com PID={proc.pid}")
                    target_exe = os.path.join(install_dir, "Mathools.exe")
                    vbs_path = _create_script_post_install_launcher(proc.pid, target_exe, exe_dir)
                    if vbs_path and os.path.exists(vbs_path):
                        wscript = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "System32", "wscript.exe")
                        subprocess.Popen(
                            [wscript, "/nologo", vbs_path],
                            creationflags=0x00000008 | 0x08000000,
                            stdin=subprocess.DEVNULL,
                            stdout=subprocess.DEVNULL,
                            stderr=subprocess.DEVNULL,
                            close_fds=True,
                        )
                        _write_error_log(f"[script] Launcher pos-instalacao iniciado: {vbs_path}")
                    else:
                        _write_error_log("[script] AVISO: launcher pos-instalacao nao foi criado.")
                    return True
                except Exception as e:
                    import traceback as _tb
                    msg = f"ERRO ao iniciar instalador em modo script: {type(e).__name__}: {e}\n{_tb.format_exc()}"
                    logger.error(f"[Updater] {msg}")
                    _write_error_log(msg)
                    return False

            logger.info("[Updater] [!] Rodando em script .py — substituição ignorada.")
            logger.info(f"[Updater] Nova versão baixada em: {new_exe}")
            return True

        # ────────────────────────────────────────────────────────────────────
        # 2. SAIR DA _MEI (garante que o processo não está dentro da pasta temp)
        # ────────────────────────────────────────────────────────────────────
        try:
            cwd_antes = os.getcwd()
            logger.debug(f"[Updater][chdir] CWD antes: {cwd_antes}")
            os.chdir(exe_dir)
            logger.debug(f"[Updater][chdir] CWD depois: {os.getcwd()} ✓")
            _write_error_log(f"[chdir] {cwd_antes} → {os.getcwd()}")
        except Exception as e:
            logger.warning(f"[Updater][chdir] Não foi possível mudar diretório: {e}")
            _write_error_log(f"[chdir] ERRO: {e}")

        # ────────────────────────────────────────────────────────────────────
        # 3. CRIAR E LANÇAR INTERMEDIÁRIO (VBS invisível → BAT)
        # ────────────────────────────────────────────────────────────────────
        try:
            # Monitora o PID do próprio app.
            # Em build onedir, getppid() pode ser explorer.exe (não encerra), travando update.
            app_pid = os.getpid()
            logger.info(f"[Updater] [2/2] Preparando intermediário de lançamento...")
            logger.debug(f"[Updater] PID monitorado do app atual: {app_pid}")

            if is_installer_package:
                vbs_path = _create_installer_launcher(new_exe, current_exe, app_pid)
            else:
                vbs_path = _create_bat_launcher(current_exe, old_exe, app_pid)
            if not vbs_path:
                logger.error("[Updater] Falha ao criar intermediário de lançamento.")
                return False

            if not os.path.exists(vbs_path) or not os.path.isfile(vbs_path):
                logger.error(f"[Updater] VBS inválido ou não criado: {vbs_path}")
                return False

            logger.info(f"[Updater] Lançando intermediário invisível...")
            logger.debug(f"[Updater] VBS validado: {vbs_path}")

            wscript = os.path.join(
                os.environ.get("SystemRoot", "C:\\Windows"), "System32", "wscript.exe"
            )
            logger.debug(f"[Updater] Executando: {wscript} /nologo {vbs_path}")

            # wscript.exe lança o VBS sem NENHUMA janela visível.
            # DETACHED_PROCESS garante que sobrevive ao fechamento deste processo.
            logger.debug(f"[Updater][launcher] wscript.exe: {wscript} | exists: {os.path.exists(wscript)}")
            _write_error_log(f"[launcher] Lançando: {wscript} /nologo {vbs_path}")

            proc = subprocess.Popen(
                [wscript, "/nologo", vbs_path],
                creationflags=0x00000008 | 0x08000000,  # DETACHED_PROCESS | CREATE_NO_WINDOW
                stdin=subprocess.DEVNULL,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                close_fds=True,
            )
            logger.debug(f"[Updater][launcher] wscript.exe lançado com PID={proc.pid}")
            _write_error_log(f"[launcher] wscript PID={proc.pid} lançado com sucesso")

            # Aguarda 1s e verifica se o processo ainda está vivo (sanity check)
            import time as _time
            _time.sleep(1)
            poll = proc.poll()
            if poll is not None:
                logger.warning(f"[Updater][launcher] wscript.exe encerrou imediatamente (returncode={poll}) — possível falha no VBS")
                _write_error_log(f"[launcher] AVISO: wscript encerrou rápido demais (rc={poll})")
            else:
                logger.debug(f"[Updater][launcher] wscript.exe ainda rodando após 1s ✓")
                _write_error_log(f"[launcher] wscript ainda vivo após 1s ✓")

        except Exception as e:
            import traceback
            msg = f"ERRO ao lançar intermediário: {type(e).__name__}: {e}\nTraceback: {traceback.format_exc()}"
            logger.error(f"[Updater] {msg}")
            _write_error_log(msg)
            return False

        # ────────────────────────────────────────────────────────────────────
        # 4. RETORNAR (não fechar processo — launcher_gui.py cuida disso)
        # .bat monitorará o PID pai e abrirá o novo EXE quando ele fechar.
        # ────────────────────────────────────────────────────────────────────
        import datetime
        ts_final = datetime.datetime.now().isoformat()
        logger.info(f"[Updater] ≡≡≡ ATUALIZAÇÃO PREPARADA — {ts_final} ≡≡≡")
        logger.info("[Updater] Feche o aplicativo para completar a atualização.")
        _write_error_log(f"[fim] download_and_apply retornou True às {ts_final}")
        return True


# ─────────────────────────────────────────────────────────────────────────────
#  API PÚBLICA
# ─────────────────────────────────────────────────────────────────────────────

_checker = UpdateChecker()


def check_update_async(callback: Callable[[Optional[UpdateInfo]], None]) -> threading.Thread:
    """
    Verifica atualização em thread separada.
    
    Args:
        callback: função chamada com UpdateInfo quando verificação terminar
    
    Returns:
        threading.Thread da verificação
    """
    return _checker.check_async(callback)


def apply_update(info: UpdateInfo, progress_cb: Optional[Callable[[int, int], None]] = None) -> bool:
    """
    Aplica atualização disponível.
    
    Fluxo:
    1. Baixa novo EXE
    2. Substitui arquivo atual
    3. Cria .bat intermediário
    4. Lança .bat
    5. Retorna True
    
    launcher_gui.py deve chamar page.window_destroy() após receber True.
    .bat monitora o processo e abre novo EXE quando o antigo fechar.
    
    Args:
        info: UpdateInfo com has_update=True
        progress_cb: callable(bytes_downloaded, total_bytes) para progresso
    
    Returns:
        True se sucesso, False se erro
    """
    return _checker.download_and_apply(info, progress_cb)




















































