"""
Utilitário de Retry Universal para Requisições HTTP
Implementa exponential backoff + jitter com configuração centralizada
"""

import time
import logging
import requests
from typing import Callable, Optional, Any, Type, Tuple
from functools import wraps
from config_loader import config

logger = logging.getLogger(__name__)


class RetryConfig:
    """Configuração de retry carregada do config.toml."""
    
    def __init__(self):
        self.enabled = config.get_bool('retry_policy.enabled', True)
        self.max_retries = config.get_int('retry_policy.max_retries', 3)
        self.initial_delay = config.get_float('retry_policy.initial_delay_sec', 1.0)
        self.max_delay = config.get_float('retry_policy.max_delay_sec', 60.0)
        self.exponential_base = config.get_float('retry_policy.exponential_base', 2.0)
        self.jitter = config.get_bool('retry_policy.jitter', True)
        self.retry_status_codes = config.get_list(
            'http.retry_status_codes',
            [408, 429, 500, 502, 503, 504]
        )


class RetryableException(Exception):
    """Exceção que indica retry é apropriado."""
    pass


def retry_on_exception(
    exception_types: Tuple[Type[Exception], ...] = None,
    max_attempts: Optional[int] = None,
    backoff_factor: Optional[float] = None,
) -> Callable:
    """
    Decorator para retry automático com exponential backoff.
    
    Args:
        exception_types: Tuple de exceções que devem disparar retry
        max_attempts: Número máximo de tentativas (usa config se None)
        backoff_factor: Fator multiplicador de delay (usa config se None)
    
    Exemplo:
        @retry_on_exception(exception_types=(requests.RequestException,))
        def fetch_data(url):
            return requests.get(url, timeout=10)
    """
    
    cfg = RetryConfig()
    
    if exception_types is None:
        exception_types = (requests.RequestException, TimeoutError)
    if max_attempts is None:
        max_attempts = cfg.max_retries + 1
    if backoff_factor is None:
        backoff_factor = cfg.exponential_base
    
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs) -> Any:
            if not cfg.enabled:
                return func(*args, **kwargs)
            
            attempt = 0
            last_exception = None
            
            while attempt < max_attempts:
                try:
                    return func(*args, **kwargs)
                except exception_types as e:
                    attempt += 1
                    last_exception = e
                    
                    if attempt >= max_attempts:
                        logger.error(
                            f"Falha permanente em {func.__name__} "
                            f"após {attempt} tentativas: {e}"
                        )
                        raise
                    
                    delay = _calculate_delay(
                        attempt - 1,
                        cfg.initial_delay,
                        cfg.max_delay,
                        backoff_factor,
                        cfg.jitter
                    )
                    
                    logger.warning(
                        f"Tentativa {attempt}/{max_attempts-1} falhou em "
                        f"{func.__name__}: {e}. "
                        f"Aguardando {delay:.2f}s antes de retry..."
                    )
                    time.sleep(delay)
            
            raise last_exception
        
        return wrapper
    return decorator


def http_request_with_retry(
    method: str,
    url: str,
    **kwargs
) -> requests.Response:
    """
    Faz requisição HTTP com retry automático baseado em status code.
    
    Args:
        method: GET, POST, PUT, DELETE, etc
        url: URL para fazer requisição
        **kwargs: argumentos adicionais para requests (headers, data, etc)
    
    Returns:
        requests.Response
    
    Raises:
        requests.HTTPError se status code não for sucesso após retries
    
    Exemplo:
        resp = http_request_with_retry('GET', 'https://example.com/api')
    """
    
    cfg = RetryConfig()
    
    if not cfg.enabled:
        resp = requests.request(method, url, **kwargs)
        resp.raise_for_status()
        return resp
    
    timeout = kwargs.get('timeout', config.get_int('http.request_timeout', 30))
    kwargs['timeout'] = timeout
    
    attempt = 0
    last_exception = None
    
    while attempt <= cfg.max_retries:
        try:
            resp = requests.request(method, url, **kwargs)
            
            if resp.status_code not in cfg.retry_status_codes:
                resp.raise_for_status()
                return resp
            
            # Status code requer retry
            if attempt >= cfg.max_retries:
                logger.error(
                    f"HTTP {resp.status_code} em {method} {url} "
                    f"após {attempt + 1} tentativas"
                )
                resp.raise_for_status()
                return resp
            
            attempt += 1
            delay = _calculate_delay(
                attempt - 1,
                cfg.initial_delay,
                cfg.max_delay,
                cfg.exponential_base,
                cfg.jitter
            )
            
            logger.warning(
                f"HTTP {resp.status_code} em {method} {url}. "
                f"Tentativa {attempt}/{cfg.max_retries}. "
                f"Aguardando {delay:.2f}s..."
            )
            time.sleep(delay)
        
        except requests.RequestException as e:
            last_exception = e
            attempt += 1
            
            if attempt > cfg.max_retries:
                logger.error(
                    f"Falha permanente em {method} {url} "
                    f"após {attempt} tentativas: {e}"
                )
                raise
            
            delay = _calculate_delay(
                attempt - 1,
                cfg.initial_delay,
                cfg.max_delay,
                cfg.exponential_base,
                cfg.jitter
            )
            
            logger.warning(
                f"Erro em {method} {url}: {e}. "
                f"Tentativa {attempt}/{cfg.max_retries}. "
                f"Aguardando {delay:.2f}s..."
            )
            time.sleep(delay)
    
    if last_exception:
        raise last_exception
    raise RuntimeError("Máximo de tentativas atingido")


def _calculate_delay(
    attempt: int,
    initial_delay: float,
    max_delay: float,
    exponential_base: float,
    use_jitter: bool
) -> float:
    """
    Calcula delay com exponential backoff + jitter opcional.
    
    Fórmula:
        delay = min(initial_delay * (exponential_base ^ attempt), max_delay)
        se jitter: delay = delay * random(0.5, 1.0)
    """
    import random
    
    delay = min(
        initial_delay * (exponential_base ** attempt),
        max_delay
    )
    
    if use_jitter:
        delay *= random.uniform(0.5, 1.0)
    
    return delay


class HTTPAdapter(requests.adapters.HTTPAdapter):
    """
    Adapter customizado para requests que implementa retry automático.
    
    Uso:
        session = requests.Session()
        adapter = HTTPAdapter()
        session.mount('http://', adapter)
        session.mount('https://', adapter)
        resp = session.get('https://example.com')
    """
    
    def __init__(self, max_retries=None, **kwargs):
        self.cfg = RetryConfig()
        from urllib3.util.retry import Retry as Urllib3Retry
        
        max_retries_obj = Urllib3Retry(
            total=self.cfg.max_retries,
            backoff_factor=self.cfg.exponential_base,
            status_forcelist=self.cfg.retry_status_codes,
            allowed_methods=['GET', 'POST', 'PUT', 'DELETE', 'HEAD'],
        )
        
        super().__init__(max_retries=max_retries_obj, **kwargs)


def create_session_with_retry() -> requests.Session:
    """
    Cria uma sessão requests com retry automático configurado.
    
    Returns:
        requests.Session com HTTPAdapter de retry
    
    Exemplo:
        session = create_session_with_retry()
        resp = session.get('https://centrosul.waterfy.net/api/...')
    """
    session = requests.Session()
    adapter = HTTPAdapter()
    session.mount('http://', adapter)
    session.mount('https://', adapter)
    return session
