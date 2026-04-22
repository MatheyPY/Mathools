"""
Circuit Breaker Pattern
Proteção contra falhas cascata em chamadas a APIs/RPC
Implementação: CLOSED → OPEN → HALF_OPEN → CLOSED
"""

import time
import logging
import threading
from enum import Enum
from typing import Callable, Any, Optional, TypeVar, Generic
from config_loader import config

logger = logging.getLogger(__name__)

T = TypeVar('T')


class CircuitState(Enum):
    """Estados possíveis do circuit breaker."""
    CLOSED = "CLOSED"          # Tudo funcionando, requisições passam
    OPEN = "OPEN"              # Falhas detectadas, requisições bloqueadas
    HALF_OPEN = "HALF_OPEN"    # Testando se serviço recuperou


class CircuitBreakerConfig:
    """Configuração do circuit breaker."""
    
    def __init__(self):
        self.enabled = config.get_bool('circuit_breaker.enabled', True)
        self.failure_threshold = config.get_int(
            'circuit_breaker.failure_threshold', 5
        )
        self.success_threshold = config.get_int(
            'circuit_breaker.success_threshold', 2
        )
        self.timeout_sec = config.get_int(
            'circuit_breaker.timeout_sec', 60
        )
        self.half_open_timeout_sec = config.get_int(
            'circuit_breaker.half_open_timeout_sec', 30
        )


class CircuitBreaker:
    """
    Circuit breaker genérico para proteção de APIs.
    
    Uso:
        breaker = CircuitBreaker(
            name="waterfy_api",
            config=CircuitBreakerConfig()
        )
        
        @breaker.handle_exceptions(
            exceptions=(RequestException,),
            fallback=lambda: {"error": "Service unavailable"}
        )
        def fetch_data():
            return requests.get("https://api.example.com/data").json()
        
        result = fetch_data()  # Protected call
    """
    
    def __init__(self, name: str, config: CircuitBreakerConfig = None):
        self.name = name
        self.cfg = config or CircuitBreakerConfig()
        
        # Estado
        self.state = CircuitState.CLOSED
        self.failure_count = 0
        self.success_count = 0
        self.last_failure_time = None
        self.last_state_change = time.time()
        
        # Thread safety
        self._lock = threading.RLock()
        
        logger.info(
            f"Circuit breaker '{name}' inicializado: "
            f"threshold={self.cfg.failure_threshold}, "
            f"timeout={self.cfg.timeout_sec}s"
        )
    
    def call(self, func: Callable[..., T], *args, **kwargs) -> T:
        """
        Executa função com proteção de circuit breaker.
        
        Args:
            func: Função a executar
            *args, **kwargs: Argumentos da função
        
        Returns:
            Resultado da função ou exceção
        
        Raises:
            CircuitBreakerOpenException se circuit está OPEN
        """
        
        if not self.cfg.enabled:
            return func(*args, **kwargs)
        
        with self._lock:
            if self.state == CircuitState.OPEN:
                if self._should_attempt_reset():
                    self.state = CircuitState.HALF_OPEN
                    self.success_count = 0
                    logger.info(
                        f"Circuit breaker '{self.name}' → HALF_OPEN "
                        "(testando recuperação)"
                    )
                else:
                    raise CircuitBreakerOpenException(
                        f"Circuit breaker '{self.name}' está OPEN. "
                        f"Próxima tentativa em {self._time_until_reset():.1f}s"
                    )
        
        try:
            result = func(*args, **kwargs)
            self._record_success()
            return result
        
        except Exception as e:
            self._record_failure()
            raise
    
    def handle_exceptions(
        self,
        exceptions: tuple = (Exception,),
        fallback: Optional[Callable] = None
    ) -> Callable:
        """
        Decorator para proteção de função com circuit breaker.
        
        Args:
            exceptions: Tuple de exceções que ativam circuit breaker
            fallback: Função alternativa se circuit abrir
        
        Exemplo:
            @breaker.handle_exceptions(
                exceptions=(RequestException,),
                fallback=lambda: mock_data
            )
            def fetch_api():
                return requests.get(...).json()
        """
        
        def decorator(func: Callable[..., T]) -> Callable[..., T]:
            def wrapper(*args, **kwargs) -> T:
                try:
                    return self.call(func, *args, **kwargs)
                except CircuitBreakerOpenException:
                    if fallback:
                        logger.warning(
                            f"Circuit breaker '{self.name}' OPEN, "
                            f"usando fallback para {func.__name__}"
                        )
                        return fallback()
                    raise
            return wrapper
        
        return decorator
    
    def _record_success(self):
        """Registra sucesso e ajusta contadores."""
        
        with self._lock:
            self.failure_count = 0
            
            if self.state == CircuitState.HALF_OPEN:
                self.success_count += 1
                
                if self.success_count >= self.cfg.success_threshold:
                    self.state = CircuitState.CLOSED
                    self.success_count = 0
                    logger.info(f"Circuit breaker '{self.name}' → CLOSED (recuperado)")
            
            else:
                self.success_count = 0
    
    def _record_failure(self):
        """Registra falha e ajusta contadores."""
        
        with self._lock:
            self.failure_count += 1
            self.last_failure_time = time.time()
            
            if self.state == CircuitState.HALF_OPEN:
                self.state = CircuitState.OPEN
                self.last_state_change = time.time()
                logger.warning(
                    f"Circuit breaker '{self.name}' → OPEN "
                    f"(falha em teste de recuperação)"
                )
            
            elif (self.state == CircuitState.CLOSED and
                  self.failure_count >= self.cfg.failure_threshold):
                self.state = CircuitState.OPEN
                self.last_state_change = time.time()
                logger.warning(
                    f"Circuit breaker '{self.name}' → OPEN "
                    f"(limiar de falhas: {self.failure_count}/{self.cfg.failure_threshold})"
                )
    
    def _should_attempt_reset(self) -> bool:
        """Verifica se tempo de timeout passou para tentar reset."""
        elapsed = time.time() - self.last_state_change
        return elapsed >= self.cfg.timeout_sec
    
    def _time_until_reset(self) -> float:
        """Retorna tempo em segundos até próxima tentativa de reset."""
        elapsed = time.time() - self.last_state_change
        return max(0, self.cfg.timeout_sec - elapsed)
    
    @property
    def status(self) -> str:
        """Retorna status legível do circuit breaker."""
        return (
            f"{self.name}: {self.state.value} | "
            f"Falhas: {self.failure_count}/{self.cfg.failure_threshold} | "
            f"Sucessos: {self.success_count}/{self.cfg.success_threshold}"
        )
    
    def reset(self):
        """Reseta circuit breaker manualmente."""
        with self._lock:
            self.state = CircuitState.CLOSED
            self.failure_count = 0
            self.success_count = 0
            self.last_failure_time = None
            logger.info(f"Circuit breaker '{self.name}' resetado manualmente")


class CircuitBreakerOpenException(Exception):
    """Exceção quando circuit breaker está OPEN."""
    pass


class CircuitBreakerManager:
    """
    Gerenciador centralizado de circuit breakers.
    
    Uso:
        manager = CircuitBreakerManager()
        waterfy_cb = manager.get_breaker("waterfy_api")
        
        @waterfy_cb.handle_exceptions(exceptions=(RequestException,))
        def fetch_data():
            return requests.get(...).json()
    """
    
    def __init__(self):
        self.breakers: dict[str, CircuitBreaker] = {}
        self._lock = threading.RLock()
    
    def get_breaker(self, name: str) -> CircuitBreaker:
        """Obtém ou cria circuit breaker por nome."""
        
        with self._lock:
            if name not in self.breakers:
                cfg = CircuitBreakerConfig()
                self.breakers[name] = CircuitBreaker(name, cfg)
            
            return self.breakers[name]
    
    def status_report(self) -> str:
        """Retorna relatório de status de todos os breakers."""
        
        with self._lock:
            if not self.breakers:
                return "Nenhum circuit breaker registrado"
            
            report = "📊 Circuit Breakers Status:\n"
            report += "=" * 70 + "\n"
            
            for breaker in self.breakers.values():
                report += f"  {breaker.status}\n"
            
            return report
    
    def reset_all(self):
        """Reseta todos os circuit breakers."""
        
        with self._lock:
            for breaker in self.breakers.values():
                breaker.reset()
            logger.info("Todos os circuit breakers resetados")


# Singleton global
circuit_breaker_manager = CircuitBreakerManager()


def get_circuit_breaker(service_name: str) -> CircuitBreaker:
    """Obtém circuit breaker para serviço específico."""
    return circuit_breaker_manager.get_breaker(service_name)


if __name__ == "__main__":
    # Teste
    import random
    
    cb = CircuitBreaker("test_api")
    
    def flaky_api_call():
        """Simula API intermitente."""
        if random.random() > 0.7:
            raise Exception("API Error")
        return {"status": "ok"}
    
    # 10 tentativas
    for i in range(10):
        try:
            result = cb.call(flaky_api_call)
            print(f"✓ {i+1}: {result}")
        except CircuitBreakerOpenException as e:
            print(f"✗ {i+1}: {e}")
        except Exception as e:
            print(f"⚠ {i+1}: {e}")
        
        print(f"   Estado: {cb.status}")
        time.sleep(1)
