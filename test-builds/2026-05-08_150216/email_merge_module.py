"""
Módulo de Mala Direta Interativa (Mail Merge)
Permite criar templates editáveis com placeholders dinâmicos para disparo massivo de emails.
"""

import re
import pandas as pd
from pathlib import Path
from typing import Optional, Dict, List, Tuple
from dataclasses import dataclass, field

# Importar configurações centralizadas de cidades para evitar duplicação
from city_configs import get_city_config, CompanyConfig


@dataclass
class MergeField:
    """Representa um campo de merge {{Nome}} no template."""
    name: str
    column_index: int = -1  # Índice da coluna no CSV


@dataclass
class MergeTemplate:
    """Template de email/documento com placeholders."""
    name: str
    subject: str
    body: str
    merge_fields: List[MergeField] = field(default_factory=list)
    created_at: str = ""
    updated_at: str = ""

    def get_all_tokens(self) -> List[str]:
        """Extrai todos os tokens {{...}} do template."""
        pattern = r"{{\s*([^{}]+?)\s*}}"
        subject_tokens = re.findall(pattern, self.subject)
        body_tokens = re.findall(pattern, self.body)
        return list(set(subject_tokens + body_tokens))


class MergeDataSource:
    """Gerencia carregamento e acesso aos dados do CSV."""
    
    def __init__(self):
        self.file_path: Optional[Path] = None
        self.dataframe: Optional[pd.DataFrame] = None
        self.columns: List[str] = []
        self.available_sheets: List[str] = []
    
    def get_excel_sheets(self, file_path: str) -> List[str]:
        """
        Retorna lista de sheets disponíveis em um arquivo Excel.
        Retorna lista vazia se não for Excel ou se houver erro.
        """
        try:
            path = Path(file_path)
            if path.suffix.lower() not in (".xlsx", ".xls", ".xlsm"):
                return []
            
            excel_file = pd.ExcelFile(path)
            sheets = excel_file.sheet_names
            self.available_sheets = list(sheets)
            return list(sheets)
        except Exception:
            return []
    
    def select_sheet(self, sheet_name: str) -> Tuple[bool, str]:
        """
        Carrega dados de uma sheet específica em um arquivo Excel.
        Retorna (sucesso, mensagem).
        """
        if not self.file_path:
            return False, "Nenhum arquivo carregado"
        
        if self.file_path.suffix.lower() not in (".xlsx", ".xls", ".xlsm"):
            return False, "Arquivo atual não é um Excel"
        
        try:
            df = pd.read_excel(
                self.file_path,
                sheet_name=sheet_name,
                dtype=str,
                keep_default_na=False
            )
            self.dataframe = df
            self.columns = list(df.columns)
            return True, f"Sheet '{sheet_name}' carregado: {len(df)} registros"
        except Exception as ex:
            return False, f"Erro ao carregar sheet: {str(ex)}"
    
    def load_csv(self, file_path: str) -> Tuple[bool, str]:
        """
        Carrega um CSV ou XLSX e retorna (sucesso, mensagem).
        Tenta múltiplas codificações e delimitadores.
        Para XLSX multi-abas: registra o file_path e carrega a PRIMEIRA sheet
        como padrão inicial. A UI deve chamar select_sheet() logo depois para
        garantir que os dados corretos estejam carregados.
        """
        try:
            path = Path(file_path)
            if not path.exists():
                return False, f"Arquivo não encontrado: {file_path}"
            
            # Se é XLSX, usa pd.read_excel diretamente
            if path.suffix.lower() in (".xlsx", ".xls", ".xlsm"):
                try:
                    # Registra o path ANTES de carregar para que select_sheet()
                    # funcione imediatamente após este método retornar.
                    self.file_path = path

                    # Descobre as abas disponíveis
                    excel_file = pd.ExcelFile(path)
                    self.available_sheets = list(excel_file.sheet_names)
                    primeira_sheet = self.available_sheets[0] if self.available_sheets else None

                    # Carrega explicitamente a primeira aba (evita ambiguidade do pandas)
                    df = pd.read_excel(
                        path,
                        sheet_name=primeira_sheet,
                        dtype=str,
                        keep_default_na=False,
                    )
                    self.dataframe = df
                    self.columns = list(df.columns)
                    sheet_info = f" [aba: {primeira_sheet}]" if primeira_sheet else ""
                    return True, f"Excel carregado: {len(df)} registros, {len(self.columns)} coluna(s){sheet_info}"
                except Exception as ex:
                    return False, f"Erro ao carregar Excel: {str(ex)}"
            
            # Para CSV, tenta múltiplas codificações e delimitadores
            delimitadores = [",", ";", "\t", "|"]
            
            for encoding in ("utf-8-sig", "utf-8", "latin-1", "iso-8859-1"):
                for delimiter in delimitadores:
                    try:
                        df = pd.read_csv(
                            path,
                            delimiter=delimiter,
                            dtype=str,
                            encoding=encoding,
                            keep_default_na=False,
                        )
                        # Valida se conseguiu ler corretamente
                        if len(df.columns) > 0 and len(df) > 0:
                            self.file_path = path
                            self.dataframe = df
                            self.columns = list(df.columns)
                            return True, f"CSV carregado: {len(df)} registros, {len(self.columns)} coluna(s) (delim='{delimiter}')"
                    except (UnicodeDecodeError, pd.errors.ParserError):
                        continue
            
            return False, "Não foi possível ler os dados. Salve sua planilha novamente como 'CSV UTF-8 (Delimitado por vírgulas)' ou como arquivo Excel padrão (.xlsx)."
            
        except Exception as ex:
            if "BadZipFile" in type(ex).__name__:
                return False, "O arquivo Excel parece estar corrompido. Tente abri-lo no seu computador e salvá-lo novamente."
            return False, "Houve um problema inesperado ao tentar ler a planilha. Verifique se ela não está protegida por senha."
    
    def get_preview(self, limit: int = 5) -> List[Dict]:
        """Retorna prévia dos dados (primeiras N linhas)."""
        if self.dataframe is None:
            return []
        
        preview = []
        for _, row in self.dataframe.head(limit).iterrows():
            preview.append(dict(row))
        return preview
    
    def iter_rows(self):
        """Itera sobre todas as linhas dos dados."""
        if self.dataframe is None:
            return
        
        for _, row in self.dataframe.iterrows():
            yield dict(row)
    
    def get_row_count(self) -> int:
        """Retorna total de registros."""
        if self.dataframe is None:
            return 0
        return len(self.dataframe)
    
    @property
    def data(self):
        """Alias para dataframe (para compatibilidade com testes e UI)."""
        return self.dataframe


# Diretório padrão de templates HTML (mesmo pacote do módulo)
TEMPLATE_DIR: Path = Path(__file__).resolve().parent / "email_templates"

# (id, rótulo exibido, nome do arquivo em email_templates/)
BUILTIN_EMAIL_TEMPLATES: List[Tuple[str, str, str]] = [
    ("itapoa_lt30", "NOTIFICAÇÃO DE DEBITOS (-30 dias)", "itapoa_notificacao_debitos.html"),
    ("itapoa_ge30", "NOTIFICAÇÃO DE DEBITOS (+30 DIAS)", "itapoa_notificacao_debitos_30dias_ou_mais.html"),
    ("notif_alto_consumo", "NOTIFICAÇÃO DE ALTO CONSUMO", "notificacao_alto_consumo.html"),
    ("notif_hidrometro_dificil", "NOTIFICAÇÃO - ACESSO DIFICIL AO HIDRÔMETRO", "notificacao_hidrometro_acesso_dificil.html"),
]


def list_builtin_email_template_choices() -> List[Tuple[str, str]]:
    """Lista templates embutidos como [(id, rótulo), ...]."""
    return [(tid, label) for tid, label, _ in BUILTIN_EMAIL_TEMPLATES]


def builtin_template_filename(template_id: str) -> Optional[str]:
    for tid, _, fn in BUILTIN_EMAIL_TEMPLATES:
        if tid == template_id:
            return fn
    return None


def load_template_html_file(path: str) -> Tuple[str, str]:
    """
    Lê um arquivo .html (UTF-8).
    Retorna (conteúdo, mensagem_erro). conteúdo vazio se falhar.
    """
    try:
        p = Path(path)
        if not p.is_file():
            return "", f"Arquivo não encontrado: {path}"
        text = p.read_text(encoding="utf-8")
        if not text.strip():
            return "", "Arquivo HTML vazio"
        return text, ""
    except Exception as ex:
        return "", str(ex)


def load_builtin_email_template(template_id: str) -> Tuple[str, str]:
    """Carrega template embutido por id. Retorna (html, erro)."""
    fn = builtin_template_filename(template_id)
    if not fn:
        return "", f"Template desconhecido: {template_id}"
    
    import os
    import sys
    
    if getattr(sys, "frozen", False):
        roots = [Path(sys._MEIPASS) / "email_templates", Path(sys.executable).parent / "email_templates"]
    else:
        roots = [TEMPLATE_DIR, Path.cwd() / "email_templates"]
        
    for root in roots:
        p = root / fn
        if p.is_file():
            return load_template_html_file(str(p))
            
    return load_template_html_file(str(TEMPLATE_DIR / fn))


def is_probably_html_document(text: str) -> bool:
    """Heurística para enviar como HTML no Outlook."""
    if not text or not text.strip():
        return False
    s = text.lstrip()[:800].lower()
    return "<html" in s or "<!doctype" in s or ("<body" in s and "<" in s)


def _column_keys_upper(columns: List[str]) -> set:
    return {str(c).strip().upper() for c in columns}


def _row_dict_upper_keys(row_data: Dict) -> Dict[str, str]:
    """Índice por nome de coluna em maiúsculas (strip) -> valor string."""
    return {str(k).strip().upper(): str(v) for k, v in row_data.items()}


class MergeEngine:
    """Motor de processamento de mail merge."""
    
    TOKEN_PATTERN = re.compile(r"{{\s*([^{}]+?)\s*}}")
    
    def __init__(self):
        self.data_source = MergeDataSource()
        self.template: Optional[MergeTemplate] = None
        self.processed_records: List[Dict] = []
        self.email_column: str = "EMAIL"

    def _resolve_email_for_row(self, row_data: Dict) -> str:
        """Obtém e-mail da linha usando self.email_column (nome de coluna, case-insensitive)."""
        row_up = _row_dict_upper_keys(row_data)
        c = str(self.email_column).strip().upper()
        if c in row_up and row_up[c].strip():
            return row_up[c].strip()
        for alt in ("EMAIL", "E-MAIL", "MAIL"):
            if alt in row_up and row_up[alt].strip():
                return row_up[alt].strip()
        return ""
    
    def extract_tokens_from_text(self, text: str) -> List[str]:
        """Extrai todos os {{tokens}} do texto."""
        matches = self.TOKEN_PATTERN.findall(text)
        return list(set(matches))  # Remove duplicatas
    
    def validate_template(self) -> Tuple[bool, str]:
        """Valida se o template está configurado corretamente."""
        if not self.template:
            return False, "Template não carregado."
        
        if not self.template.subject.strip():
            return False, "Assunto do email está vazio."
        
        if not self.template.body.strip():
            return False, "Corpo do email está vazio."
        
        tokens = self.template.get_all_tokens()
        if not tokens:
            return False, "Template não possui placeholders {{...}}. Adicione pelo menos um campo!"
        
        col_up = _column_keys_upper(list(self.data_source.columns))
        
        # Variáveis virtuais injetadas por empresa/unidade (previne erro na validação)
        virtual_cols = {
            "EMPRESA_NOME", "EMPRESA_NOME_CURTO", "EMPRESA_CNPJ", "EMPRESA_ENDERECO",
            "EMPRESA_0800", "EMPRESA_0800_LINK", "EMPRESA_SITE", "EMPRESA_WHATSAPP",
            "EMPRESA_WHATSAPP_LINK", "EMPRESA_MUNICIPIO", "EMPRESA_APP", "EMPRESA_APP_LINK",
            "COR_PRIMARIA", "COR_SECUNDARIA", "COR_TERCIARIA", "EMPRESA_LOGO_B64", "EMPRESA_HORARIO"
        }
        col_up.update(virtual_cols)
        
        missing_cols: List[str] = []
        for token in tokens:
            if str(token).strip().upper() not in col_up:
                missing_cols.append(token)
        
        if missing_cols:
            cols_str = ", ".join(missing_cols)
            return False, f"Colunas não encontradas no CSV (comparação ignora maiúsculas): {cols_str}"
        
        return True, "Template válido!"
    
    def _resolve_company_info(self, row_up: Dict[str, str]) -> Dict[str, str]:
        """Resolve dados dinâmicos da empresa baseado na coluna CIDADE_IMOVEL da planilha."""
        cidade_row = row_up.get("CIDADE_IMOVEL", row_up.get("CIDADE", ""))
        cidade = cidade_row.strip()
        
        # Busca configuração centralizada de city_configs.py
        company_config = get_city_config(cidade)
        
        if company_config is None:
            # Fallback para configuração padrão se cidade não encontrada
            company_config = CompanyConfig(
                name="SANEAMENTO",
                cnpj="",
                address="",
                phone="",
                whatsapp="",
                website="",
                app="",
                app_link="https://play.google.com/store/apps/details?id=com.waterfy.agenciavirtual.centrosul",
                hours="de segunda a sexta, das 8h às 17h",
                city=cidade,
                state="",
            )
        
        # Converte para dict de merge compatível com templates
        info = company_config.to_merge_dict()
        
        return info


    def render_record(self, row_data: Dict) -> Dict[str, str]:
        """
        Renderiza um registro individual (substitui tokens).
        Retorna dicionário com subject e body com dados substituídos.
        """
        if not self.template:
            return {"subject": "", "body": ""}
        
        row_up = _row_dict_upper_keys(row_data)

        # Injeta variáveis da empresa baseadas na cidade
        company_info = self._resolve_company_info(row_up)
        for k, v in company_info.items():
            if k not in row_up:
                row_up[k] = v

        def replace_token(match):
            token = match.group(1).strip()
            key = token.upper()
            if key in row_up:
                return row_up[key]
            return f"[{token} não encontrado]"
        
        rendered_subject = self.TOKEN_PATTERN.sub(replace_token, self.template.subject)
        rendered_body = self.TOKEN_PATTERN.sub(replace_token, self.template.body)
        
        return {"subject": rendered_subject, "body": rendered_body}
    
    def process_batch(
        self,
        max_records: Optional[int] = None,
        email_column: Optional[str] = None,
    ) -> Tuple[int, List[str]]:
        """
        Processa todos os registros gerando emails renderizados.
        Retorna (total_processado, lista_de_erros).
        email_column: nome da coluna do destinatário (ex.: EMAIL); opcional, usa o já configurado.
        """
        if not self.template:
            return 0, ["Template não configurado"]
        if email_column:
            self.email_column = str(email_column).strip() or "EMAIL"

        valid, msg = self.validate_template()
        if not valid:
            return 0, [msg]
        
        self.processed_records = []
        errors = []
        processed = 0
        
        try:
            for i, row_data in enumerate(self.data_source.iter_rows()):
                if max_records and i >= max_records:
                    break
                
                try:
                    rendered = self.render_record(row_data)
                    em = self._resolve_email_for_row(row_data)
                    self.processed_records.append({
                        "row_index": i,
                        "subject": rendered["subject"],
                        "body": rendered["body"],
                        "email": em,
                        "row_data": row_data,
                    })
                    processed += 1
                except Exception as ex:
                    errors.append(f"Linha {i+1}: {str(ex)}")
        
        except Exception as ex:
            errors.append(f"Erro ao processar batch: {str(ex)}")
        
        return processed, errors
    
    def get_preview_record(self, record_index: int = 0) -> Optional[Dict]:
        """Retorna prévia de um registro processado."""
        if record_index < 0 or record_index >= len(self.processed_records):
            return None
        return self.processed_records[record_index]
    
    def export_json(self, output_path: str) -> Tuple[bool, str]:
        """Exporta todos os registros processados em JSON."""
        try:
            import json
            data = {
                "template": {
                    "name": self.template.name if self.template else "",
                    "subject": self.template.subject if self.template else "",
                    "body": self.template.body if self.template else "",
                },
                "total_records": len(self.processed_records),
                "records": self.processed_records,
            }
            with open(output_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            return True, f"Exportado para: {output_path}"
        except Exception as ex:
            return False, f"Erro ao exportar: {str(ex)}"


def create_default_template() -> MergeTemplate:
    """Cria um template padrão de exemplo."""
    return MergeTemplate(
        name="Comunicado Padrão",
        subject="Comunicado para {{Nome}}",
        body=(
            "Prezado(a) {{Nome}},\n\n"
            "Endereço registrado: {{Endereco}}\n"
            "Valor: {{Valor}}\n\n"
            "Favor entran em contato conosco para mais informações.\n\n"
            "Atenciosamente,\nEquipe Mathools"
        ),
    )