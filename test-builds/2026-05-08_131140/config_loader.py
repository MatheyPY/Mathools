"""
Carregador de Configuração Centralizado
Lê config.toml e fornece acesso a todas as configurações
"""

import os
import sys
import json
import shutil  # <--- Adicionado para a mágica de copiar o arquivo
from pathlib import Path
from typing import Any, Dict, Optional

# Tentar importar tomli (compatível com Python < 3.11)
try:
    import tomllib  # Python 3.11+
except ImportError:
    try:
        import tomli as tomllib  # Fallback
    except ImportError:
        tomllib = None


def _get_base_dir() -> Path:
    """Retorna diretório base (pasta do .exe ou script)."""
    if getattr(sys, 'frozen', False):
        # PyInstaller: apontar para a pasta REAL onde o .exe está rodando
        return Path(sys.executable).parent
    return Path(__file__).parent


class ConfigLoader:
    """Carregador centralizado de configuração."""
    
    _instance = None
    _config: Dict[str, Any] = {}
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._load_config()
        return cls._instance
    
    def _load_config(self):
        """Carrega config.toml na primeira instância."""
        base_dir = _get_base_dir()
        config_path = base_dir / "config.toml"
        
        # =========================================================
        # MÁGICA DE RECUPERAÇÃO:
        # Se o arquivo não existir do lado de fora, tenta puxar 
        # o padrão empacotado no PyInstaller e extrair ele para a pasta!
        # =========================================================
        if getattr(sys, 'frozen', False) and not config_path.exists():
            internal_path = Path(sys._MEIPASS) / "config.toml"
            if internal_path.exists():
                try:
                    shutil.copy2(internal_path, config_path)
                except Exception:
                    # Se der erro de permissão ao salvar, lê o interno mesmo
                    config_path = internal_path
        
        if not config_path.exists():
            raise FileNotFoundError(
                f"config.toml não encontrado em {config_path}. "
                "Por favor, crie o arquivo de configuração."
            )
        
        if tomllib is None:
            raise ImportError(
                "Biblioteca 'tomli' (ou Python 3.11+) necessária para ler TOML. "
                "Instale: pip install tomli"
            )
        
        try:
            with open(config_path, 'rb') as f:
                self._config = tomllib.load(f)
        except Exception as e:
            raise ValueError(f"Erro ao ler config.toml: {e}")
    
    def get(self, path: str, default: Any = None) -> Any:
        """
        Obtém valor de configuração usando notação de ponto.
        Exemplo: config.get('http.retry_attempts') → 3
        """
        keys = path.split('.')
        value = self._config
        
        try:
            for key in keys:
                value = value[key]
            return value
        except (KeyError, TypeError):
            if default is not None:
                return default
            raise KeyError(f"Chave de configuração não encontrada: {path}")
    
    def get_bool(self, path: str, default: bool = False) -> bool:
        """Obtém valor booleano."""
        val = self.get(path, default)
        return bool(val)
    
    def get_int(self, path: str, default: int = 0) -> int:
        """Obtém valor inteiro."""
        val = self.get(path, default)
        return int(val)
    
    def get_float(self, path: str, default: float = 0.0) -> float:
        """Obtém valor float."""
        val = self.get(path, default)
        return float(val)
    
    def get_str(self, path: str, default: str = "") -> str:
        """Obtém valor string."""
        val = self.get(path, default)
        return str(val)
    
    def get_list(self, path: str, default: list = None) -> list:
        """Obtém valor lista."""
        if default is None:
            default = []
        val = self.get(path, default)
        return list(val) if val else default
    
    def reload(self):
        """Recarrega configuração (útil em testes)."""
        self._config = {}
        self._load_config()
    
    @property
    def config_dict(self) -> Dict[str, Any]:
        """Retorna dict inteiro de configuração."""
        return self._config.copy()
    
    def validate_required_keys(self, required_keys: list):
        """Valida se chaves obrigatórias existem."""
        missing = []
        for key in required_keys:
            try:
                self.get(key)
            except KeyError:
                missing.append(key)
        
        if missing:
            raise ValueError(
                f"Chaves de configuração obrigatórias ausentes: {missing}"
            )


# Singleton global
config = ConfigLoader()


# Aliases para acesso fácil (estilo Django settings)
def get_setting(path: str, default: Any = None) -> Any:
    """Acesso direto a uma configuração."""
    return config.get(path, default)


if __name__ == "__main__":
    # Teste rápido
    cfg = ConfigLoader()
    print("🔧 Configurações carregadas:")
    print(f"  - App: {cfg.get_str('core.app_version')}")
    print(f"  - Debug: {cfg.get_bool('core.debug_mode')}")
    print(f"  - Retry habilitado: {cfg.get_bool('http.enable_retry')}")
    print(f"  - Circuit breaker: {cfg.get_bool('circuit_breaker.enabled')}")
    print(f"  - Timezone: {cfg.get_str('core.timezone')}")