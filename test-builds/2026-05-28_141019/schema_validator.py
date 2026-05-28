"""
Validador de Schema para CSVs
Valida estrutura e tipos de dados antes do processamento
Conformidade com especificação de macro Waterfy
"""

import pandas as pd
import logging
from typing import Dict, List, Optional, Any
from enum import Enum
from config_loader import config

logger = logging.getLogger(__name__)


class DataType(Enum):
    """Tipos de dados esperados em CSV."""
    STRING = "string"
    INTEGER = "integer"
    FLOAT = "float"
    DATE = "date"
    BOOLEAN = "boolean"
    NUMERIC = "numeric"  # int ou float


class CSVSchema:
    """Define esperado de um CSV."""
    
    def __init__(
        self,
        name: str,
        required_columns: List[str],
        optional_columns: List[str] = None,
        column_types: Dict[str, DataType] = None,
    ):
        self.name = name
        self.required_columns = required_columns
        self.optional_columns = optional_columns or []
        self.column_types = column_types or {}
    
    def validate_dataframe(self, df: pd.DataFrame) -> tuple[bool, List[str]]:
        """
        Valida DataFrame contra schema.
        
        Returns:
            (is_valid, list_of_errors)
        """
        errors = []
        
        # 1. Validar colunas obrigatórias
        missing = [c for c in self.required_columns if c not in df.columns]
        if missing:
            errors.append(f"Colunas obrigatórias faltando: {missing}")
        
        # 2. Validar tipos de dados
        for col, expected_type in self.column_types.items():
            if col not in df.columns:
                continue
            
            try:
                self._validate_column_type(df[col], col, expected_type)
            except ValueError as e:
                errors.append(str(e))
        
        return (len(errors) == 0, errors)
    
    def _validate_column_type(
        self,
        series: pd.Series,
        col_name: str,
        expected_type: DataType
    ):
        """Valida tipo de coluna individual."""
        
        # Ignorar NaN
        non_null = series.dropna()
        
        if expected_type == DataType.STRING:
            if not all(isinstance(x, str) for x in non_null):
                raise ValueError(
                    f"Coluna '{col_name}' esperada STRING, "
                    f"encontrado {non_null.dtype}"
                )
        
        elif expected_type == DataType.INTEGER:
            try:
                pd.to_numeric(non_null, downcast='integer')
            except (ValueError, TypeError):
                raise ValueError(
                    f"Coluna '{col_name}' esperada INTEGER, "
                    f"contém valores não-numéricos"
                )
        
        elif expected_type == DataType.FLOAT:
            try:
                pd.to_numeric(non_null, downcast='float')
            except (ValueError, TypeError):
                raise ValueError(
                    f"Coluna '{col_name}' esperada FLOAT, "
                    f"contém valores não-numéricos"
                )
        
        elif expected_type == DataType.NUMERIC:
            try:
                pd.to_numeric(non_null)
            except (ValueError, TypeError):
                raise ValueError(
                    f"Coluna '{col_name}' esperada NUMERIC, "
                    f"contém valores não-numéricos"
                )
        
        elif expected_type == DataType.DATE:
            try:
                pd.to_datetime(non_null)
            except (ValueError, TypeError):
                raise ValueError(
                    f"Coluna '{col_name}' esperada DATE, "
                    f"não conseguiu fazer parse"
                )
        
        elif expected_type == DataType.BOOLEAN:
            valid_bool = {True, False, 1, 0, 'S', 'N', 'Y', 'N', 'sim', 'não'}
            invalid = [x for x in non_null if x not in valid_bool]
            if invalid:
                raise ValueError(
                    f"Coluna '{col_name}' esperada BOOLEAN, "
                    f"encontrado valores inválidos: {invalid[:3]}"
                )


# Schemas padrão para cada macro
SCHEMA_8117 = CSVSchema(
    name="8117 - Boletim Diário de Arrecadação",
    required_columns=[
        "DATA", "ARRECADADOR", "MOVIMENTO", "QTD_FATURAS",
        "VALOR_ARRECADADO"
    ],
    optional_columns=["DESCRICAO", "OBSERVACOES"],
    column_types={
        "DATA": DataType.DATE,
        "QTD_FATURAS": DataType.INTEGER,
        "VALOR_ARRECADADO": DataType.FLOAT,
    }
)

SCHEMA_8121 = CSVSchema(
    name="8121 - Arrecadação x Faturamento Diário/Acumulado",
    required_columns=[
        "DATA", "FATURAMENTO", "ARRECADACAO", "TAXA_ARRECADACAO",
        "FATURAMENTO_ACUM", "ARRECADACAO_ACUM"
    ],
    optional_columns=["META", "DIVERGENCIA"],
    column_types={
        "DATA": DataType.DATE,
        "FATURAMENTO": DataType.FLOAT,
        "ARRECADACAO": DataType.FLOAT,
        "TAXA_ARRECADACAO": DataType.FLOAT,
    }
)

SCHEMA_8091 = CSVSchema(
    name="8091 - Faturamento Arrecadação Diário/Acumulado",
    required_columns=[
        "MES", "FATURAMENTO", "ARRECADACAO", "EM_ATRASO",
        "CONECTADAS", "ATIVAS"
    ],
    optional_columns=["CIDADE", "CATEGORIA_CLIENTE"],
    column_types={
        "MES": DataType.DATE,
        "FATURAMENTO": DataType.FLOAT,
        "ARRECADACAO": DataType.FLOAT,
        "EM_ATRASO": DataType.FLOAT,
        "CONECTADAS": DataType.INTEGER,
        "ATIVAS": DataType.INTEGER,
    }
)

SCHEMA_EXP50012 = CSVSchema(
    name="50012 - Relatório Analítico (LAT/LONG + Medição)",
    required_columns=[
        "MATRICULA", "CLIENTE", "ENDERECO", "LATITUDE", "LONGITUDE",
        "LEITURA_ATUAL", "LEITURA_ANTERIOR", "CONSUMO", "VALOR"
    ],
    optional_columns=["BAIRRO", "CATEGORIA", "SITUACAO", "EMAIL", "TELEFONE"],
    column_types={
        "MATRICULA": DataType.STRING,
        "LATITUDE": DataType.FLOAT,
        "LONGITUDE": DataType.FLOAT,
        "LEITURA_ATUAL": DataType.NUMERIC,
        "LEITURA_ANTERIOR": DataType.NUMERIC,
        "CONSUMO": DataType.NUMERIC,
        "VALOR": DataType.FLOAT,
    }
)

SCHEMA_8104 = CSVSchema(
    name="8104 - Base para Cobrança/WhatsApp",
    required_columns=[
        "MATRICULA", "CLIENTE_DA_FATURA", "EMAIL", "TEL_CEL",
        "ENDERECO", "DIVIDA_DO_INQUILINO_ATUAL", "QTDE_DIAS_EM_ATRASO"
    ],
    optional_columns=["CATEGORIA", "BAIRRO", "SITUACAO"],
    column_types={
        "MATRICULA": DataType.STRING,
        "DIVIDA_DO_INQUILINO_ATUAL": DataType.FLOAT,
        "QTDE_DIAS_EM_ATRASO": DataType.INTEGER,
    }
)

# Mapa de macros para schemas
SCHEMA_MAP = {
    "8117": SCHEMA_8117,
    "8121": SCHEMA_8121,
    "8091": SCHEMA_8091,
    "50012": SCHEMA_EXP50012,
    "8104": SCHEMA_8104,
    "macro_8104": SCHEMA_8104,
}


def validate_csv_file(
    filepath: str,
    macro_id: str = None,
    strict: bool = False
) -> Dict[str, Any]:
    """
    Valida arquivo CSV contra schema apropriado.
    
    Args:
        filepath: Caminho do arquivo CSV
        macro_id: ID da macro (8117, 8121, 8091, 50012, 8104)
        strict: Se True, falha em aviso; se False, apenas log
    
    Returns:
        {
            'valid': bool,
            'errors': List[str],
            'warnings': List[str],
            'dataframe': pd.DataFrame (se válido)
        }
    
    Exemplo:
        result = validate_csv_file('dados.csv', macro_id='8091')
        if result['valid']:
            df = result['dataframe']
    """
    
    result = {
        'valid': False,
        'errors': [],
        'warnings': [],
        'dataframe': None,
        'macro_id': macro_id,
    }
    
    try:
        # 1. Carregar CSV
        df = pd.read_csv(filepath, encoding='utf-8')
    except Exception as e:
        result['errors'].append(f"Erro ao ler CSV: {e}")
        logger.error(f"Falha ao ler {filepath}: {e}")
        return result
    
    # 2. Detectar macro_id automaticamente se necessário
    if not macro_id:
        macro_id = _detect_macro_id(df)
        if not macro_id:
            result['warnings'].append(
                "Macro ID não detectado automaticamente. "
                "Validação básica apenas."
            )
    
    # 3. Validar contra schema
    schema = SCHEMA_MAP.get(macro_id)
    
    if schema:
        is_valid, schema_errors = schema.validate_dataframe(df)
        result['errors'].extend(schema_errors)
    else:
        result['warnings'].append(
            f"Schema não definido para macro {macro_id}. "
            f"Validação básica apenas."
        )
    
    # 4. Validar linhas vazias e duplicatas
    if df.empty:
        result['errors'].append("CSV está vazio")
    else:
        null_pct = (df.isnull().sum().sum() / (len(df) * len(df.columns)) * 100)
        if null_pct > 50:
            result['warnings'].append(
                f"CSV tem {null_pct:.1f}% de valores nulos"
            )
        
        duplicates = df.duplicated().sum()
        if duplicates > 0:
            result['warnings'].append(
                f"CSV contém {duplicates} linhas duplicadas"
            )
    
    # 5. Resultado final
    result['valid'] = len(result['errors']) == 0
    
    if result['valid']:
        result['dataframe'] = df
        logger.info(
            f"CSV validado com sucesso: {filepath} "
            f"({len(df)} linhas, {len(df.columns)} colunas)"
        )
    else:
        logger.error(
            f"CSV inválido: {filepath}. Erros: {result['errors']}"
        )
    
    if result['warnings']:
        logger.warning(f"Avisos em {filepath}: {result['warnings']}")
    
    return result


def _detect_macro_id(df: pd.DataFrame) -> Optional[str]:
    """Detecta macro ID baseado em colunas presentes."""
    
    cols = set(df.columns)
    
    # Heurísticas de detecção
    if {'LATITUDE', 'LONGITUDE'}.issubset(cols):
        return "50012"
    elif {'TAXA_ARRECADACAO'}.issubset(cols):
        return "8121"
    elif {'ARRECADADOR'}.issubset(cols):
        return "8117"
    elif {'TEL_CEL', 'DIVIDA_DO_INQUILINO_ATUAL'}.issubset(cols):
        return "8104"
    elif {'FATURAMENTO_ACUM', 'ARRECADACAO_ACUM'}.issubset(cols):
        return "8091"
    
    return None
