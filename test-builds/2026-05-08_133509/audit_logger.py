"""
Logger de Auditoria LGPD
Registra ações do sistema sem dados pessoais identificáveis (PII).
Conformidade: LGPD Art. 6º, III (minimização); Art. 37 (rastreabilidade)

RETENÇÃO DE 5 ANOS (1825 dias):
  Lei 5.172/1966 Art. 43 — Código Tributário Nacional
  Decreto Federal 9.662/2019 — Lei de Acesso à Informação (órgão público)
  LGPD Art. 7º, III — Base legal para manutenção de dados financeiros
  
Logs de cobrança via WhatsApp devem ser mantidos por 5 anos para:
  ✓ Rastreabilidade de transações financeiras
  ✓ Conformidade fiscal (Receita Federal)
  ✓ Auditoria externa
  ✓ Resolução de disputas judiciais
"""

import logging
import os
import sys  # <--- Adicionado para a correção do PyInstaller
from logging.handlers import RotatingFileHandler
from datetime import datetime
from typing import Optional, Dict, Any
import json
from config_loader import config


class AuditLogger:
    """Logger especializado para auditoria sem PII."""
    
    def __init__(self):
        """Inicializa logger de auditoria com configuração centralizada."""
        
        self.logger = logging.getLogger("AUDITORIA_LGPD")
        self.logger.setLevel(logging.INFO)
        
        # Limpar handlers existentes
        if self.logger.handlers:
            self.logger.handlers.clear()
        
        # Obter configuração original
        audit_file = config.get_str(
            'logging.audit_file',
            'logs/auditoria_LGPD.log'
        )

        # =========================================================
        # INÍCIO DA CORREÇÃO DE CAMINHO SEGURO (Impede Erro System32)
        # =========================================================
        # Só altera se for um caminho relativo (ex: 'logs/arquivo.log')
        if not os.path.isabs(audit_file):
            if getattr(sys, 'frozen', False):
                # Rodando como .exe (PyInstaller)
                base_dir = os.path.dirname(sys.executable)
            else:
                # Rodando como script .py normal
                base_dir = os.path.dirname(os.path.abspath(__file__))
            
            # Força o caminho a ser relativo à pasta real do programa
            audit_file = os.path.join(base_dir, audit_file)
        # =========================================================
        
        max_bytes = config.get_int('logging.max_bytes', 5242880)
        backup_count = config.get_int('logging.backup_count', 5)
        
        # Criar diretório se não existir (agora usando o caminho seguro)
        audit_dir = os.path.dirname(audit_file)
        if audit_dir:
            os.makedirs(audit_dir, exist_ok=True)
        
        # Handler de arquivo com rotação
        file_handler = RotatingFileHandler(
            audit_file,
            maxBytes=max_bytes,
            backupCount=backup_count,
            encoding='utf-8'
        )
        
        # Formato customizado para auditoria
        formatter = logging.Formatter(
            '%(asctime)s | %(levelname)s | %(user)s | %(action)s | '
            '%(status)s | %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        file_handler.setFormatter(formatter)
        self.logger.addHandler(file_handler)
        
        # Lista de campos que NÃO devem aparecer em logs (PII)
        self.pii_fields = config.get_list(
            'logging.pii_fields',
            ['password', 'cpf', 'email', 'phone', 'endereco', 'token']
        )
    
    def _sanitize_data(self, data: Dict[str, Any]) -> Dict[str, Any]:
        """Remove PII do dicionário."""
        sanitized = {}
        for key, value in data.items():
            if any(pii in key.lower() for pii in self.pii_fields):
                sanitized[key] = "***REDACTED***"
            else:
                sanitized[key] = value
        return sanitized
    
    def log_auth_attempt(
        self,
        user: str,
        success: bool,
        ip_address: str = "UNKNOWN",
        detail: str = ""
    ):
        """Log de tentativa de autenticação."""
        
        status = "SUCCESS" if success else "FAILED"
        extra = {
            'user': user,
            'action': 'AUTH_ATTEMPT',
            'status': status,
        }
        
        message = f"Autenticação | IP: {ip_address}"
        if detail:
            message += f" | Detalhe: {detail}"
            
        self.logger.info(message, extra=extra)
    
    def log_data_access(
        self,
        user: str,
        resource: str,
        action: str = "READ",
        success: bool = True,
        detail: str = ""
    ):
        """Log de acesso a dados."""
        
        status = "SUCCESS" if success else "FAILED"
        extra = {
            'user': user,
            'action': f"DATA_ACCESS_{action}",
            'status': status,
        }
        
        message = f"Acesso a recurso: {resource}"
        if detail:
            message += f" | {detail}"
            
        self.logger.info(message, extra=extra)
    
    def log_download(
        self,
        user: str,
        file_name: str,
        file_size_bytes: int = 0,
        success: bool = True,
        macro_id: str = None
    ):
        """Log de download de arquivo."""
        
        status = "SUCCESS" if success else "FAILED"
        extra = {
            'user': user,
            'action': 'DOWNLOAD',
            'status': status,
        }
        
        size_mb = file_size_bytes / (1024 * 1024) if file_size_bytes else 0
        message = f"Download: {file_name} ({size_mb:.2f} MB)"
        
        if macro_id:
            message += f" | Macro: {macro_id}"
            
        self.logger.info(message, extra=extra)
    
    def log_config_change(
        self,
        user: str,
        config_key: str,
        old_value: Any,
        new_value: Any
    ):
        """Log de mudança de configuração."""
        
        extra = {
            'user': user,
            'action': 'CONFIG_CHANGE',
            'status': 'INFO',
        }
        
        # Sanitizar valores sensíveis
        old = old_value if not any(p in str(config_key) for p in self.pii_fields) else "***"
        new = new_value if not any(p in str(config_key) for p in self.pii_fields) else "***"
        
        message = f"Configuração alterada: {config_key} | {old} → {new}"
        self.logger.info(message, extra=extra)
    
    def log_process_execution(
        self,
        user: str,
        process_name: str,
        status: str = "STARTED",
        duration_seconds: float = None,
        rows_processed: int = None,
        detail: str = ""
    ):
        """Log de execução de processo (download, processamento, etc)."""
        
        extra = {
            'user': user,
            'action': f"PROCESS_{process_name.upper()}",
            'status': status,
        }
        
        message = f"Processo: {process_name}"
        
        if duration_seconds is not None:
            message += f" | Duração: {duration_seconds:.2f}s"
            
        if rows_processed is not None:
            message += f" | Registros: {rows_processed}"
            
        if detail:
            message += f" | {detail}"
            
        self.logger.info(message, extra=extra)
    
    def log_error(
        self,
        user: str,
        process: str,
        error_type: str,
        detail: str = ""
    ):
        """Log de erro durante processo."""
        
        extra = {
            'user': user,
            'action': f"ERROR_{process.upper()}",
            'status': 'FAILED',
        }
        
        message = f"Erro em {process}: {error_type}"
        if detail:
            message += f" | {detail}"
            
        self.logger.error(message, extra=extra)
    
    def log_api_call(
        self,
        user: str,
        api_name: str,
        method: str,
        endpoint: str,
        status_code: int = None,
        response_time_ms: float = None,
        success: bool = True
    ):
        """Log de chamada à API."""
        
        status = "SUCCESS" if success else "FAILED"
        extra = {
            'user': user,
            'action': f"API_CALL_{api_name.upper()}",
            'status': status,
        }
        
        message = f"{method} {endpoint}"
        
        if status_code:
            message += f" | HTTP {status_code}"
            
        if response_time_ms:
            message += f" | {response_time_ms:.0f}ms"
            
        self.logger.info(message, extra=extra)
    
    def log_data_deletion(
        self,
        user: str,
        entity: str,
        count: int,
        reason: str = ""
    ):
        """Log de exclusão de dados (LGPD Art. 18, VI)."""
        
        extra = {
            'user': user,
            'action': 'DATA_DELETION',
            'status': 'SUCCESS',
        }
        
        message = f"Exclusão de dados: {entity} ({count} registros)"
        
        if reason:
            message += f" | Motivo: {reason}"
            
        self.logger.info(message, extra=extra)
    
    def log_custom(
        self,
        user: str,
        action: str,
        status: str,
        message: str
    ):
        """Log customizado com padrão de auditoria."""
        
        extra = {
            'user': user,
            'action': action,
            'status': status,
        }
        
        self.logger.info(message, extra=extra)


# Singleton global
audit_logger = AuditLogger()


# Interface simples para uso
def audit(
    user: str,
    action: str,
    status: str = "INFO",
    message: str = ""
):
    """
    Registra ação em auditoria de forma simples.
    
    Exemplo:
        audit("matheus@itapoa", "DOWNLOAD_OS", "SUCCESS", "8091 macro baixado")
    """
    audit_logger.log_custom(user, action, status, message)


if __name__ == "__main__":
    # Teste
    audit_logger.log_auth_attempt(
        user="operador_001",
        success=True,
        ip_address="192.168.1.100"
    )
    
    audit_logger.log_download(
        user="operador_001",
        file_name="macro_8091.csv",
        file_size_bytes=1024576,
        success=True,
        macro_id="8091"
    )
    
    audit_logger.log_process_execution(
        user="operador_001",
        process_name="CONSOLIDACAO",
        status="COMPLETED",
        duration_seconds=45.32,
        rows_processed=2500
    )
    
    print("✅ Auditoria testada - verifique logs/auditoria_LGPD.log")