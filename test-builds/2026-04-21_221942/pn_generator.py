"""
Gerador de PN (Painel de Negócios)
Função especializada para baixar 8117 + 8121 + 8091 e gerar PN em HTML
"""

import os
import logging
import threading
import queue
import calendar
from datetime import datetime
from typing import Optional, Dict, Tuple

try:
    import waterfy_engine as eng
    import dashboard_html as gen
    from audit_logger import audit_logger
    from timezone_utils import tz_manager
    from config_loader import config
    from cache_manager import get_cache_manager
except ImportError as e:
    print(f"[AVISO] Módulo não encontrado: {e}")

logger = logging.getLogger(__name__)


def gerar_pn_com_download(
    usuario: str,
    senha: str,
    data_inicial: str,
    data_final: str,
    output_dir: str = "Dashboard_Waterfy",
    log_fn=None,
    base_dir: Optional[str] = None
) -> Tuple[bool, str, Optional[str]]:
    """
    Genera PN baixando 8117 + 8121 do Waterfy.
    
    Args:
        usuario: Email/usuário Waterfy
        senha: Senha Waterfy
        data_inicial: DD/MM/YYYY
        data_final: DD/MM/YYYY
        output_dir: Diretório de saída (padrão: Dashboard_Waterfy)
        log_fn: Função de callback para logs (opcional)
    
    Returns:
        (sucesso: bool, mensagem: str, caminho_arquivo: str|None)
    """
    
    try:
        def _meses_no_intervalo(di: str, df: str):
            """Converte intervalo DD/MM/YYYY para lista de meses MM/YYYY (inclusive)."""
            dt_i = datetime.strptime(di, "%d/%m/%Y")
            dt_f = datetime.strptime(df, "%d/%m/%Y")
            cur = datetime(dt_i.year, dt_i.month, 1)
            fim = datetime(dt_f.year, dt_f.month, 1)
            meses = []
            while cur <= fim:
                meses.append(cur.strftime("%m/%Y"))
                nm = cur.month + 1
                na = cur.year + (1 if nm > 12 else 0)
                nm = 1 if nm > 12 else nm
                cur = datetime(na, nm, 1)
            return meses

        # 1. Valida datas
        try:
            d_ini = datetime.strptime(data_inicial, "%d/%m/%Y")
            d_fim = datetime.strptime(data_final, "%d/%m/%Y")
        except ValueError as e:
            return False, f"Formato de data inválido: {e}", None
        
        if d_ini > d_fim:
            return False, "Data inicial não pode ser maior que data final", None
        
        # 2. Cria diretório de saída se não existir
        os.makedirs(output_dir, exist_ok=True)
        
        # 3. Cache manager para cache mensal de dados 8091
        cache_mgr = get_cache_manager(output_dir)
        
        # 4. Log de início
        try:
            audit_logger.log_process_execution(
                user=usuario,
                process_name="GERAR_PN_COM_DOWNLOAD",
                status="STARTED",
                detail=f"{data_inicial} → {data_final}"
            )
        except Exception:
            pass
        
        # 5. Busca macros (8117 + 8121 + 8091) via waterfy_engine ───────────────
        # Cria callback de log combinado (console + arquivo + UI)
        def _log_combined(msg):
            """Envia logs para console, arquivo e interface (se fornecida)."""
            print(msg, flush=True)
            logging.info(msg)
            if log_fn:
                try:
                    log_fn(msg)
                except Exception:
                    pass
        
        _log_combined("Buscando dados de Arrecadação e Faturamento...")
        
        try:
            # Monta periodos para ambas as macros (data_inicial, data_final)
            meses_8091 = _meses_no_intervalo(data_inicial, data_final)
            periodos = {
                '8117': (data_inicial, data_final),
                '8121': (data_inicial, data_final),
                '8091': meses_8091,
            }
            
            # Chama buscar_macros UMA VEZ com ambas
            # (internamente ele usa threading de 2 sessions!)
            dados = eng.buscar_macros(
                usuario=usuario,
                senha=senha,
                macros=['8117', '8121', '8091'],
                periodos=periodos,
                log_fn=_log_combined,
                cache_mgr=cache_mgr,
                base_dir=base_dir
            )
            
            _log_combined("Compilando dados...")
            
        except Exception as e:
            logging.error(f"[ERRO] Erro ao baixar macros: {e}")
            _log_combined(f"Erro ao processar dados: {e}")
            return False, f"Erro ao baixar macros: {e}", None
        
        # Valida se conseguiu dados
        if not dados or 'erro' in dados:
            erro = dados.get('erro', 'Macros retornaram vazio')
            return False, f"Erro ao processar macros: {erro}", None
        
        if 'macro_8117' not in dados or 'macro_8121' not in dados:
            return False, "Não conseguiu baixar as macros obrigatórias do PN (8117 e 8121)", None
        
        print("Dados processados com sucesso!")
        
        # 5. Monta dados consolidados para geração de PN ─────────────────────────
        print("Consolidando dados...")
        
        # Dados já consolidados por buscar_macros()
        # Adiciona informações de período e timestamp
        dados['periodo'] = {
            'inicial': data_inicial,
            'final': data_final
        }
        dados['cidade'] = 'ITAPOÁ'
        dados['timestamp_gerado'] = tz_manager.now_iso()
        
        # 6. Gera PN em HTML
        print("Gerando relatório...")
        
        timestamp = tz_manager.now().strftime("%Y-%m-%d_%H%M%S")
        nome_arquivo = f"PN_Financeiro_{timestamp}.html"
        caminho_saida = os.path.join(output_dir, nome_arquivo)
        
        # Chama gerador de PN
        resultado = gen.gerar_pn_html(dados, caminho_saida)
        
        if not resultado or not os.path.exists(caminho_saida):
            return False, "Erro ao gerar arquivo PN em HTML", None
        
        # 7. Log de sucesso
        tamanho_mb = os.path.getsize(caminho_saida) / (1024 * 1024)
        try:
            audit_logger.log_process_execution(
                user=usuario,
                process_name="GERAR_PN_COM_DOWNLOAD",
                status="COMPLETED",
                rows_processed=len(dados.get('macro_8117', {}).get('dados', [])),
                detail=f"PN gerado: {nome_arquivo} ({tamanho_mb:.2f} MB)"
            )
            audit_logger.log_download(
                user=usuario,
                file_name=nome_arquivo,
                file_size_bytes=os.path.getsize(caminho_saida),
                success=True,
                macro_id="PN_8117+8121+8091"
            )
        except Exception:
            pass
        
        mensagem = f" PN gerado com sucesso!\n {nome_arquivo}\n Período: {data_inicial} a {data_final}"
        return True, mensagem, caminho_saida
    
    except Exception as e:
        logger.error(f"Erro ao gerar PN: {e}", exc_info=True)
        
        # Log de erro
        try:
            audit_logger.log_error(
                user=usuario,
                process="GERAR_PN_COM_DOWNLOAD",
                error_type=type(e).__name__,
                detail=str(e)
            )
        except Exception:
            pass
        
        return False, f"Erro ao gerar PN: {str(e)}", None


def validar_credenciais_waterfy(usuario: str, senha: str) -> Tuple[bool, str]:
    """
    Valida credenciais Waterfy fazendo login teste.
    
    Returns:
        (válidas: bool, mensagem: str)
    """
    try:
        # Tenta fazer login para validar
        resultado = eng.fazer_login(usuario, senha)
        
        if resultado and 'SESSID' in resultado:
            return True, "Credenciais válidas"
        else:
            return False, "Credenciais inválidas"
    
    except Exception as e:
        return False, f"Erro ao validar credenciais: {str(e)}"
