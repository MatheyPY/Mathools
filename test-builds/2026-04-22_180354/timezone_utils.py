"""
Utilitário de Timezone
Padroniza timestamps em toda a aplicação com fuso horário correto
Conformidade: Logs e auditoria com timestamps globalmente consistentes
"""

from datetime import datetime, timedelta, timezone
from typing import Optional
import logging
from zoneinfo import ZoneInfo
from config_loader import config

logger = logging.getLogger(__name__)


class TimezoneManager:
    """Gerenciador centralizado de fusos horários."""
    
    def __init__(self):
        # Obter timezone da configuração
        tz_string = config.get_str('core.timezone', 'America/Sao_Paulo')
        
        try:
            self.app_tz = ZoneInfo(tz_string)
            self.tz_string = tz_string
            logger.info(f"Timezone inicializado: {tz_string}")
        except Exception as e:
            logger.warning(f"Erro ao usar {tz_string}: {e}. Usando UTC.")
            self.app_tz = ZoneInfo('UTC')
            self.tz_string = 'UTC'
    
    def now(self) -> datetime:
        """
        Retorna data/hora atual no timezone da aplicação.
        
        Returns:
            datetime com timezone consciente
        
        Exemplo:
            dt = tz_manager.now()  # 2026-03-24 14:30:45-03:00
        """
        return datetime.now(self.app_tz)
    
    def now_iso(self) -> str:
        """
        Retorna timestamp ISO 8601 atual.
        
        Returns:
            String ISO formatada
        
        Exemplo:
            "2026-03-24T14:30:45.123456-03:00"
        """
        return self.now().isoformat()
    
    def now_timestamp(self) -> float:
        """
        Retorna Unix timestamp do hora atual.
        
        Returns:
            Float de timestamp Unix (segundos desde epoch)
        """
        return self.now().timestamp()
    
    def format_datetime(
        self,
        dt: Optional[datetime] = None,
        format_str: str = "%d/%m/%Y %H:%M:%S"
    ) -> str:
        """
        Formata datetime com timezone da app.
        
        Args:
            dt: datetime a formatar (usa now() se None)
            format_str: formato strftime
        
        Returns:
            String formatada
        
        Exemplo:
            tz_manager.format_datetime(format_str="%d de %B de %Y às %H:%M")
            # "24 de March de 2026 às 14:30"
        """
        if dt is None:
            dt = self.now()
        elif dt.tzinfo is None:
            # Se datetime é naive, assumir UTC e converter
            dt = dt.replace(tzinfo=timezone.utc).astimezone(self.app_tz)
        else:
            # Se tem timezone, converter para app_tz
            dt = dt.astimezone(self.app_tz)
        
        return dt.strftime(format_str)
    
    def parse_datetime(
        self,
        date_string: str,
        format_str: str = "%d/%m/%Y %H:%M:%S"
    ) -> datetime:
        """
        Parse string de data em timezone da app.
        
        Args:
            date_string: String de data
            format_str: Format esperado
        
        Returns:
            datetime consciente de timezone
        
        Exemplo:
            dt = tz_manager.parse_datetime("24/03/2026 14:30")
        """
        dt_naive = datetime.strptime(date_string, format_str)
        # Assumir que input está em timezone da app
        dt_aware = dt_naive.replace(tzinfo=self.app_tz)
        return dt_aware
    
    def to_utc(self, dt: datetime) -> datetime:
        """Converte datetime para UTC."""
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=self.app_tz)
        return dt.astimezone(timezone.utc)
    
    def from_utc(self, dt: datetime) -> datetime:
        """Converte datetime UTC para timezone da app."""
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.astimezone(self.app_tz)
    
    def datetime_with_offset(self) -> tuple[datetime, str]:
        """
        Retorna (datetime, offset) da hora atual.
        
        Returns:
            (datetime, string de offset como "-03:00")
        
        Exemplo:
            dt, offset = tz_manager.datetime_with_offset()
            print(f"{dt} UTC{offset}")  # 2026-03-24 14:30:45 UTC-03:00
        """
        dt = self.now()
        offset = dt.strftime("%z")
        # Formatar offset para -03:00
        offset_fmt = f"{offset[:3]}:{offset[3:]}"
        return dt, offset_fmt
    
    def days_since(self, dt: datetime) -> int:
        """Calcula dias desde data específica."""
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=self.app_tz)
        
        delta = self.now() - dt
        return delta.days
    
    def timestamp_to_datetime(self, timestamp: float) -> datetime:
        """Converte Unix timestamp para datetime com timezone."""
        dt = datetime.fromtimestamp(timestamp, tz=timezone.utc)
        return dt.astimezone(self.app_tz)
    
    @property
    def utc_offset(self) -> str:
        """Retorna offset UTC da app (ex: -03:00)."""
        dt = self.now()
        offset = dt.strftime("%z")
        return f"{offset[:3]}:{offset[3:]}"
    
    @property
    def tz_name(self) -> str:
        """Retorna nome do timezone (ex: BRST, BRT)."""
        return self.now().strftime("%Z")


# Singleton global
tz_manager = TimezoneManager()


# Funções de conveniência
def now() -> datetime:
    """Retorna agora com timezone correto."""
    return tz_manager.now()


def now_iso() -> str:
    """Retorna timestamp ISO atual."""
    return tz_manager.now_iso()


def format_time(dt: Optional[datetime] = None, fmt: str = "%d/%m/%Y %H:%M:%S") -> str:
    """Formata datetime com timezone."""
    return tz_manager.format_datetime(dt, fmt)


def parse_time(date_str: str, fmt: str = "%d/%m/%Y %H:%M:%S") -> datetime:
    """Parse string de data."""
    return tz_manager.parse_datetime(date_str, fmt)


class TimestampMixin:
    """
    Mixin para classes que precisam de timestamps com timezone.
    
    Uso:
        class LogEntry(TimestampMixin):
            def __init__(self):
                self.created_at = self.now()
                self.updated_at = self.now()
    """
    
    def now(self) -> datetime:
        """Retorna agora no timezone app."""
        return tz_manager.now()
    
    def now_iso(self) -> str:
        """Retorna ISO timestamp."""
        return tz_manager.now_iso()
    
    def format_time(self, dt: Optional[datetime] = None) -> str:
        """Formata datetime."""
        return tz_manager.format_datetime(dt)


if __name__ == "__main__":
    # Teste
    print("🕐 Testes de Timezone Manager:")
    print(f"  Timezone app: {tz_manager.tz_string}")
    print(f"  Agora: {tz_manager.now()}")
    print(f"  ISO 8601: {tz_manager.now_iso()}")
    print(f"  Offset UTC: {tz_manager.utc_offset}")
    print(f"  Nome TZ: {tz_manager.tz_name}")
    
    # Teste parse
    parsed = tz_manager.parse_datetime("24/03/2026 10:30", "%d/%m/%Y %H:%M")
    print(f"\n  Parsed: {parsed}")
    print(f"  Em UTC: {tz_manager.to_utc(parsed)}")
    
    # Teste formatação customizada
    formatted = tz_manager.format_datetime(format_str="%A, %d de %B de %Y às %H:%M")
    print(f"\n  Formatado: {formatted}")
    
    print("\n✅ Timezone Manager funcional!")
