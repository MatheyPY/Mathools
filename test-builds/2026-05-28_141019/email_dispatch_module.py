from __future__ import annotations

import html
import re
import smtplib
import base64
import tempfile
import webbrowser
from dataclasses import dataclass, field
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from pathlib import Path
from typing import Callable, Iterable, Optional, Protocol

import flet as ft
import pandas as pd

try:
    from docx_merge_module import DocxMergeEngine, DocxDataSource
    HAS_DOCX_MERGE = True
except ImportError:
    HAS_DOCX_MERGE = False

# Importar configurações centralizadas de cidades para evitar duplicação
from city_configs import CompanyConfig, DEFAULT_COMPANIES, get_city_config

# Importar logger de auditoria LGPD
try:
    from audit_logger import audit_logger
    HAS_AUDIT_LOGGER = True
except ImportError:
    HAS_AUDIT_LOGGER = False
    audit_logger = None


def _safe_text(value: object) -> str:
    if value is None:
        return ""
    if pd.isna(value):
        return ""
    return str(value)


def _strip_html_tags_plain(html_str: str, limit: int = 8000) -> str:
    t = re.sub(r"(?is)<script.*?>.*?</script>", " ", html_str)
    t = re.sub(r"(?is)<style.*?>.*?</style>", " ", t)
    t = re.sub(r"<[^>]+>", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t[:limit] + ("…" if len(t) > limit else "")


def _flet_html_control(initial: str) -> ft.Control:
    Html = getattr(ft, "Html", None)
    if Html is not None:
        return Html(data=initial)
    return ft.Text(
        _strip_html_tags_plain(initial, 600),
        size=13,
        color="#627D98",
        selectable=True,
    )


def _set_flet_html_data(control: ft.Control, html: str) -> None:
    if hasattr(control, "data"):
        control.data = html
    elif isinstance(control, ft.Text):
        control.value = _strip_html_tags_plain(html, 12000)


def _dispatch_email_body_preview_widget(html_content: str) -> ft.Control:
    """Painel embutido com ft.Html se existir; senão abre o HTML no navegador (Flet 0.82 / Windows)."""
    Html = getattr(ft, "Html", None)
    if Html is not None:
        try:
            return ft.Container(
                bgcolor="#FFFFFF",
                border_radius=ft.BorderRadius.only(bottom_left=12, bottom_right=12),
                padding=0,
                content=ft.Column(
                    spacing=0,
                    controls=[Html(data=html_content, expand=True)],
                ),
            )
        except Exception:
            pass
    tmp = Path(tempfile.gettempdir()) / "mathtools_dispatch_email_preview.html"
    try:
        tmp.write_text(html_content, encoding="utf-8")
    except Exception:
        pass

    def _open_preview(_):
        if tmp.is_file():
            webbrowser.open(tmp.as_uri())

    return ft.Container(
        bgcolor="#FFFFFF",
        border_radius=ft.BorderRadius.only(bottom_left=12, bottom_right=12),
        padding=16,
        content=ft.Column(
            tight=True,
            spacing=8,
            controls=[
                ft.Text(
                    "Nesta versão do Flet no Windows não há painel HTML embutido. "
                    "O arquivo abaixo contém o mesmo HTML do envio — use o botão para abrir no navegador.",
                    size=12,
                    color="#627D98",
                ),
                ft.FilledButton("Abrir prévia no navegador", on_click=_open_preview),
                ft.Text(str(tmp), size=10, color="#94A3B8", selectable=True),
            ],
        ),
    )


@dataclass(slots=True)
class EmailTemplateModel:
    name: str
    subject: str
    body: str
    description: str = ""


@dataclass(slots=True)
class SmtpServerConfig:
    host: str
    port: int
    username: str
    password: str
    use_tls: bool = True
    from_name: str = "Mathools"
    from_email: Optional[str] = None


@dataclass(slots=True)
class EmailDispatchPayload:
    recipient: str
    subject: str
    html_body: str


def _sanitize_email_body_text(text: str) -> str:
    """Remove trechos indesejados e normaliza o corpo antes da renderização/envio."""
    sanitized = text.replace("(PDD: PDD 365 +).", "")
    sanitized = sanitized.replace("(PDD: PDD 365 +)", "")
    sanitized = re.sub(r"\n{3,}", "\n\n", sanitized)
    return sanitized.strip()

def _get_logo_base64(filename: Optional[str] = None) -> str:
    """
    Mantido por compatibilidade, mas a logo preferencial vem via EMPRESA_LOGO_B64
    já resolvida pelo email_merge_module (_resolve_company_info).
    Só é usado como último recurso caso custom_data não contenha a chave.
    """
    import os, sys, base64
    candidates = [filename] if filename else []
    if getattr(sys, "frozen", False):
        roots = [sys._MEIPASS, os.path.dirname(sys.executable)]
    else:
        roots = [os.path.abspath("."), os.path.dirname(os.path.abspath(__file__))]
    for root in roots:
        for name in candidates:
            if not name:
                continue
            p = os.path.join(root, name)
            if os.path.exists(p):
                try:
                    with open(p, "rb") as f:
                        b64 = base64.b64encode(f.read()).decode("ascii")
                    ext = os.path.splitext(name)[1].lower().lstrip(".")
                    mime = "jpeg" if ext in ("jpg", "jpeg") else "png"
                    return f"data:image/{mime};base64,{b64}"
                except Exception:
                    pass
    return ""


def _load_template_html(filename: str) -> str:
    """
    Carrega um template HTML do diretório email_templates.
    Retorna a string vazia se o arquivo não for encontrado.
    """
    import os, sys
    if getattr(sys, "frozen", False):
        roots = [sys._MEIPASS, os.path.dirname(sys.executable)]
    else:
        roots = [os.path.abspath("."), os.path.dirname(os.path.abspath(__file__))]
    
    for root in roots:
        template_path = os.path.join(root, "email_templates", filename)
        if os.path.exists(template_path):
            try:
                with open(template_path, "r", encoding="utf-8") as f:
                    return f.read()
            except Exception:
                pass
    return ""

class SpreadsheetDataSource:
    def __init__(self) -> None:
        self.file_path: Optional[Path] = None
        self._sheets: dict[str, pd.DataFrame] = {}
        self._active_sheet: Optional[str] = None

    @property
    def active_sheet(self) -> Optional[str]:
        return self._active_sheet

    def load_file(self, file_path: str) -> list[str]:
        path = Path(file_path)
        suffix = path.suffix.lower()
        self.file_path = path
        self._sheets.clear()
        self._active_sheet = None

        if suffix == ".csv":
            dataframe = self._read_csv(path)
            self._sheets["CSV"] = dataframe.copy(deep=True)
        elif suffix in {".xlsx", ".xlsm", ".xls"}:
            excel_file = pd.ExcelFile(path)
            for sheet_name in excel_file.sheet_names:
                dataframe = pd.read_excel(
                    path,
                    sheet_name=sheet_name,
                    dtype=object,
                    keep_default_na=False,
                )
                self._sheets[sheet_name] = dataframe.copy(deep=True)
        else:
            raise ValueError("Formato de arquivo não suportado. Use CSV ou Excel.")

        sheet_names = list(self._sheets.keys())
        if sheet_names:
            self.select_sheet(sheet_names[0])
        return sheet_names

    def select_sheet(self, sheet_name: str) -> None:
        if sheet_name not in self._sheets:
            raise KeyError(f"Aba não encontrada: {sheet_name}")
        self._active_sheet = sheet_name

    def list_columns(self) -> list[str]:
        dataframe = self.get_active_dataframe()
        return [str(column) for column in dataframe.columns]

    def preview_rows(self, limit: int = 8) -> list[dict[str, str]]:
        dataframe = self.get_active_dataframe().head(limit)
        preview: list[dict[str, str]] = []
        for _, row in dataframe.iterrows():
            preview.append({str(column): _safe_text(value) for column, value in row.items()})
        return preview

    def iter_rows(self) -> Iterable[dict[str, str]]:
        dataframe = self.get_active_dataframe()
        for _, row in dataframe.iterrows():
            yield {str(column): _safe_text(value) for column, value in row.items()}

    def get_active_dataframe(self) -> pd.DataFrame:
        if not self._active_sheet:
            raise ValueError("Nenhuma origem de dados carregada.")
        return self._sheets[self._active_sheet]

    @staticmethod
    def _read_csv(path: Path) -> pd.DataFrame:
        last_error: Optional[Exception] = None
        for encoding in ("utf-8-sig", "utf-8", "latin-1"):
            try:
                return pd.read_csv(
                    path,
                    dtype=object,
                    keep_default_na=False,
                    encoding=encoding,
                )
            except UnicodeDecodeError as exc:
                last_error = exc
        if last_error:
            raise last_error
        raise ValueError("Não foi possível ler o CSV informado.")


class TemplateCatalog:
    def __init__(self, templates: Optional[list[EmailTemplateModel]] = None) -> None:
        self._templates = templates or self._default_templates()

    def list_templates(self) -> list[EmailTemplateModel]:
        return list(self._templates)

    def get_template(self, name: str) -> EmailTemplateModel:
        for template in self._templates:
            if template.name == name:
                return template
        raise KeyError(f"Modelo não encontrado: {name}")

    @staticmethod
    def _default_templates() -> list[EmailTemplateModel]:
        return [
            EmailTemplateModel(
                name="Aviso Amigável 14 dias",
                subject="Lembrete de pendência financeira - {{Nome}}",
                description="Abordagem cordial para atrasos iniciais.",
                body=(
                    "Olá {{Nome}},\n\n"
                    "Identificamos uma pendência vinculada à matrícula {{Matrícula}} "
                    "no valor de {{Valor}} com {{Dias_Atraso}} dias de atraso.\n\n"
                    "Se o pagamento já foi realizado, por favor desconsidere esta mensagem. "
                    "Caso precise de apoio, nossa equipe está à disposição.\n\n"
                    "Atenciosamente,\nEquipe Mathools"
                ),
            ),
            EmailTemplateModel(
                name="Notificação de Protesto 30+ dias",
                subject="Notificação formal de débito - {{Nome}}",
                description="Texto mais firme para réguas avançadas.",
                body=(
                    "Prezado(a) {{Nome}},\n\n"
                    "Consta em nosso sistema débito referente à matrícula {{Matrícula}}, "
                    "no valor de {{Valor}}, em atraso há {{Dias_Atraso}} dias.\n\n"
                    "Solicitamos regularização imediata para evitar encaminhamentos "
                    "administrativos e medidas adicionais.\n\n"
                    "Atenciosamente,\nEquipe Mathools"
                ),
            ),
            EmailTemplateModel(
                name="Notificação de Alto Consumo",
                subject="Notificação de alto consumo de água — {{CIDADE_IMOVEL}}",
                description="Aviso sobre consumo de água acima da média.",
                body=_load_template_html("notificacao_alto_consumo.html"),
            ),
            EmailTemplateModel(
                name="Padronização - Hidrômetro de Difícil Acesso",
                subject="Padronização - Acesso ao Hidrômetro — {{CIDADE_IMOVEL}}",
                description="Solicitação de padronização para acesso ao hidrômetro.",
                body=_load_template_html("notificacao_hidrometro_acesso_dificil.html"),
            ),
            EmailTemplateModel(
                name="Campanha Especial de Negociação (Desenrola Brasil)",
                subject="Campanha Especial de Negociação — {{CIDADE_IMOVEL}}",
                description="Template HTML para campanha especial de negociação de débitos com condições facilitadas.",
                body=_load_template_html("campanha_desenrola_brasil.html"),
            ),
        ]


class MergeFieldEngine:
    TOKEN_PATTERN = re.compile(r"{{\s*([^{}]+?)\s*}}")

    def insert_token(self, current_text: str, column_name: str) -> str:
        token = f"{{{{{column_name}}}}}"
        separator = "" if not current_text or current_text.endswith((" ", "\n")) else " "
        return f"{current_text}{separator}{token}"

    def render(self, text: str, row_data: dict[str, str]) -> str:
        def replace(match: re.Match[str]) -> str:
            key = match.group(1).strip()
            return row_data.get(key, "")

        return self.TOKEN_PATTERN.sub(replace, text)


class CorporateEmailHtmlBuilder:
    def __init__(
        self,
        brand_name: str = "Mathools",
        accent_color: str = "#F27D16",
        heading_color: str = "#0D1F35",
        surface_color: str = "#F4F6FA",
    ) -> None:
        self.brand_name = brand_name
        self.accent_color = accent_color
        self.heading_color = heading_color
        self.surface_color = surface_color

    def build_html(self, subject: str, plain_body: str) -> str:
        paragraphs = []
        for block in plain_body.split("\n\n"):
            normalized = "<br>".join(html.escape(line) for line in block.splitlines())
            if normalized.strip():
                paragraphs.append(
                    "<p style='margin:0 0 18px;font-size:16px;line-height:1.7;color:#334155;'>"
                    f"{normalized}</p>"
                )
        body_html = "".join(paragraphs) or (
            "<p style='margin:0;font-size:16px;line-height:1.7;color:#334155;'>&nbsp;</p>"
        )
        escaped_subject = html.escape(subject)
        
        brand_html = f'''<div style="display:inline-block;padding:8px 14px;border:1px solid rgba(255,255,255,0.28);border-radius:999px;font-size:12px;letter-spacing:0.08em;text-transform:uppercase;color:#ffffff;">
              {html.escape(self.brand_name)}
            </div>'''

        return f"""\
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{escaped_subject}</title>
</head>
<body style="margin:0;padding:0;background:#E9EEF5;font-family:Segoe UI,Arial,sans-serif;">
  <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:#E9EEF5;margin:0;padding:24px 12px;">
    <tr>
      <td align="center">
        <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="max-width:760px;background:#ffffff;border-radius:24px;overflow:hidden;">
          <tr>
            <td style="padding:0;background:linear-gradient(135deg, {self.heading_color} 0%, #16375D 58%, {self.accent_color} 100%);">
              <table role="presentation" width="100%" cellspacing="0" cellpadding="0">
                <tr>
                  <td style="padding:28px 32px 18px;">
                        {brand_html}
                    <h1 style="margin:18px 0 0;font-size:30px;line-height:1.2;color:#ffffff;font-weight:700;">
                      {escaped_subject}
                    </h1>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td style="padding:32px;background:{self.surface_color};">
              <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:#ffffff;border-radius:18px;">
                <tr>
                  <td style="padding:30px 28px 12px;">
                    {body_html}
                  </td>
                </tr>
                <tr>
                  <td style="padding:0 28px 28px;">
                    <div style="height:1px;background:#E2E8F0;"></div>
                    <p style="margin:18px 0 0;font-size:13px;line-height:1.6;color:#64748B;">
                      Este e-mail foi gerado automaticamente pelo módulo corporativo de comunicação do Mathools.
                    </p>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>
"""


# As configurações de empresa são centralizadas em city_configs.py
# e importadas no topo do arquivo. Este módulo usa DEFAULT_COMPANIES apenas
# para o seletor manual de unidade na UI (city_dropdown).


class DynamicEmailHtmlBuilder:
    """Builder de HTML dinâmico que se adapta a diferentes empresas/cidades."""

    def __init__(self, company_config: Optional[CompanyConfig] = None, template_style: str = "modern"):
        self.company_config = company_config or self._default_config()
        self.template_style = template_style  # "modern", "classic", "minimal"

    def _default_config(self) -> CompanyConfig:
        """Configuração padrão neutra — a empresa real é resolvida pelo email_merge_module via CIDADE_IMOVEL."""
        return CompanyConfig(
            name="SANEAMENTO",
            cnpj="",
            address="",
            phone="",
            whatsapp="",
            website="",
            app="",
            app_link="",
            hours="de segunda a sexta, das 8h às 17h",
        )

    def set_company_config(self, config: CompanyConfig) -> None:
        """Atualiza a configuração da empresa dinamicamente."""
        self.company_config = config

    def build_html(self, subject: str, plain_body: str, custom_data: Optional[dict] = None) -> str:
        """
        Gera HTML dinâmico baseado na configuração atual.

        A logo e as cores da empresa são lidas EXCLUSIVAMENTE de custom_data,
        cujas chaves (EMPRESA_LOGO_B64, COR_PRIMARIA, COR_SECUNDARIA, etc.)
        já foram resolvidas pelo email_merge_module via CIDADE_IMOVEL.
        Não há lógica de seleção de empresa aqui.

        Args:
            subject: Assunto do email
            plain_body: Corpo do email
            custom_data: Dados da linha já mesclados (contém EMPRESA_LOGO_B64 e cores)
        """
        config = self.company_config

        if custom_data:
            row_up = {str(k).upper(): str(v) for k, v in custom_data.items()}

            city_name = ""
            for candidate in ("CIDADE_IMOVEL", "CIDADE", "MUNICIPIO"):
                if candidate in row_up and row_up[candidate].strip():
                    city_name = row_up[candidate].strip()
                    break
            if not city_name:
                for key, value in row_up.items():
                    if ("CIDADE" in key or "MUNICIPIO" in key) and value.strip():
                        city_name = value.strip()
                        break
            if city_name:
                city_config = get_city_config(city_name)
                if city_config:
                    config = city_config

            # Sobrescrever config apenas com os campos de empresa que o merge_module
            # já resolveu corretamente — sem tentar re-detectar cidade aqui.
            field_map = {
                "EMPRESA_NOME": "name",
                "EMPRESA_CNPJ": "cnpj",
                "EMPRESA_ENDERECO": "address",
                "EMPRESA_0800": "phone",
                "EMPRESA_WHATSAPP": "whatsapp",
                "EMPRESA_SITE": "website",
                "EMPRESA_APP": "app",
                "EMPRESA_HORARIO": "hours",
                "COR_PRIMARIA": "primary_color",
                "COR_SECUNDARIA": "secondary_color",
                "EMPRESA_LOGO_B64": "logo_url",
            }
            from dataclasses import replace as dc_replace
            overrides = {}
            for token_key, config_field in field_map.items():
                if token_key in row_up and row_up[token_key]:
                    overrides[config_field] = row_up[token_key]
            if overrides:
                config = dc_replace(config, **overrides)

        paragraphs = []
        for block in plain_body.split("\n\n"):
            normalized = "<br>".join(html.escape(line) for line in block.splitlines())
            if normalized.strip():
                paragraphs.append(
                    f"<p class='MsoNormal' style='text-align:justify'>"
                    f"<b>{normalized}</b><o:p></o:p></p>"
                )
        body_html = "\n".join(paragraphs) or "<p class='MsoNormal'>&nbsp;</p>"

        escaped_subject = html.escape(subject)

        if self.template_style == "modern":
            return self._build_modern_template(config, escaped_subject, body_html)
        elif self.template_style == "classic":
            return self._build_classic_template(config, escaped_subject, body_html)
        else:
            return self._build_minimal_template(config, escaped_subject, body_html)

    def _build_modern_template(self, config: CompanyConfig, subject: str, body: str) -> str:
        """Template moderno com gradiente (padrão Mathools)."""
        # config.logo_url agora pode ser um data URI completo (vindo do EMPRESA_LOGO_B64)
        # ou um nome de arquivo em disco (fallback legado).
        logo_b64 = ""
        if config.logo_url:
            if config.logo_url.startswith("data:"):
                logo_b64 = config.logo_url
            else:
                logo_b64 = _get_logo_base64(config.logo_url)
        if logo_b64:
            brand_html = f'<img src="{logo_b64}" alt="{html.escape(config.name)}" style="width:100%; max-width:100%; height:auto; display:block; border:0; outline:none; text-decoration:none; border-radius:8px;">'
        else:
            brand_html = f'''<div style="display:inline-block;padding:8px 14px;border:1px solid rgba(255,255,255,0.28);border-radius:999px;font-size:12px;letter-spacing:0.08em;text-transform:uppercase;color:#ffffff;">
              {html.escape(config.name)}
            </div>'''
        return f"""\
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{subject}</title>
</head>
<body style="margin:0;padding:0;background:#E9EEF5;font-family:Segoe UI,Arial,sans-serif;">
  <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:#E9EEF5;margin:0;padding:24px 12px;">
    <tr>
      <td align="center">
        <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="max-width:760px;background:#ffffff;border-radius:24px;overflow:hidden;">
          <tr>
            <td style="padding:0;background:linear-gradient(135deg, {config.primary_color} 0%, #16375D 58%, {config.secondary_color} 100%);">
              <table role="presentation" width="100%" cellspacing="0" cellpadding="0">
                <tr>
                  <td style="padding:28px 32px 18px;">
                        {brand_html}
                    <h1 style="margin:18px 0 0;font-size:30px;line-height:1.2;color:#ffffff;font-weight:700;">
                      {subject}
                    </h1>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td style="padding:32px;background:#F4F6FA;">
              <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:#ffffff;border-radius:18px;">
                <tr>
                  <td style="padding:30px 28px 12px;">
                    {body}
                  </td>
                </tr>
                <tr>
                  <td style="padding:0 28px 28px;">
                    <div style="height:1px;background:#E2E8F0;"></div>
                    <p style="margin:18px 0 0;font-size:13px;line-height:1.6;color:#64748B;">
                      Este e-mail foi gerado automaticamente pelo sistema corporativo.
                    </p>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>
"""

    def _build_classic_template(self, config: CompanyConfig, subject: str, body: str) -> str:
        """Template clássico estilo Word/Outlook (padrão Itapoá)."""
        return f"""\
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <meta name="Generator" content="Microsoft Word 15 (filtered medium)">
  <style>
    @font-face {{font-family:"Cambria Math";panose-1:2 4 5 3 5 4 6 3 2 4;}}
    @font-face {{font-family:Aptos;}}
    p.MsoNormal, li.MsoNormal, div.MsoNormal {{margin:0cm;font-size:12.0pt;font-family:"Aptos",sans-serif;}}
    .MsoChpDefault {{font-size:12.0pt;font-family:"Aptos",sans-serif;}}
    .MsoPapDefault {{margin-bottom:8.0pt;line-height:115%;}}
  </style>
</head>
<body lang="PT-BR" style='word-wrap:break-word'>
  <div class="WordSection1">
    <p class="MsoNormal"><b><span style='font-size:11.0pt'>{html.escape(config.name)} - {html.escape(config.cnpj)}<o:p></o:p></span></b></p>
    <p class="MsoNormal"><b><span style='font-size:11.0pt'>{html.escape(config.address)}<o:p></o:p></span></b></p>
    <p class="MsoNormal"><b><span style='font-size:11.0pt'>ATENDIMENTO: {html.escape(config.phone)}<o:p></o:p></span></b></p>
    <p class="MsoNormal"><b><span style='font-size:11.0pt'>{html.escape(config.website)}<o:p></o:p></span></b></p>
    <p class="MsoNormal"><b><o:p>&nbsp;</o:p></b></p>
    <p class="MsoNormal" align="center" style='text-align:center'>
      <b><span style='font-size:14.0pt'>{subject}<o:p></o:p></span></b>
    </p>
    {body}
    <p class="MsoNormal"><b><o:p>&nbsp;</o:p></b></p>
    <p class="MsoNormal" style='text-align:justify'><b>Para facilitar, utilize nossos canais oficiais:</b></p>
    <p class="MsoNormal" style='text-align:justify'><b>📞 {html.escape(config.phone)}</b></p>
    <p class="MsoNormal" style='text-align:justify'><b>💬 WhatsApp: {html.escape(config.whatsapp)}</b></p>
    <p class="MsoNormal" style='text-align:justify'><b>🌐 {html.escape(config.website)}</b></p>
    <p class="MsoNormal" style='text-align:justify'><b>📱 Aplicativo {html.escape(config.app)}</b></p>
    <p class="MsoNormal" style='text-align:justify'><b>• Atendimento presencial: {html.escape(config.hours)}</b></p>
    <p class="MsoNormal"><b><o:p>&nbsp;</o:p></b></p>
    <p class="MsoNormal" align="center" style='text-align:center'><b>Atenciosamente,<o:p></o:p></b></p>
    <p class="MsoNormal" align="center" style='text-align:center'><b>{html.escape(config.name)}<o:p></o:p></b></p>
  </div>
</body>
</html>
"""

    def _build_minimal_template(self, config: CompanyConfig, subject: str, body: str) -> str:
        """Template minimalista e limpo."""
        return f"""\
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <title>{subject}</title>
</head>
<body style="margin:0;padding:20px;font-family:Arial,sans-serif;background:#f5f5f5;">
  <div style="max-width:600px;margin:0 auto;background:white;padding:30px;border-radius:8px;box-shadow:0 2px 4px rgba(0,0,0,0.1);">
    <div style="border-bottom:2px solid {config.primary_color};padding-bottom:15px;margin-bottom:20px;">
      <h2 style="margin:0;color:{config.primary_color};">{html.escape(config.name)}</h2>
      <p style="margin:5px 0 0;color:#666;font-size:12px;">{html.escape(config.cnpj)}</p>
    </div>
    <h1 style="color:{config.primary_color};font-size:20px;">{subject}</h1>
    <div style="line-height:1.6;color:#333;">
      {body}
    </div>
    <div style="margin-top:30px;padding-top:20px;border-top:1px solid #eee;font-size:12px;color:#666;">
      <p style="margin:5px 0;">📞 {html.escape(config.phone)}</p>
      <p style="margin:5px 0;">🌐 {html.escape(config.website)}</p>
    </div>
  </div>
</body>
</html>
"""


class SmtpEmailSender:
    def __init__(self, config: SmtpServerConfig) -> None:
        self.config = config

    def send(self, payload: EmailDispatchPayload) -> None:
        message = MIMEMultipart("alternative")
        from_email = self.config.from_email or self.config.username
        message["Subject"] = payload.subject
        message["From"] = f"{self.config.from_name} <{from_email}>"
        message["To"] = payload.recipient
        message.attach(MIMEText(payload.html_body, "html", "utf-8"))

        with smtplib.SMTP(self.config.host, self.config.port) as server:
            if self.config.use_tls:
                server.starttls()
            server.login(self.config.username, self.config.password)
            server.sendmail(from_email, [payload.recipient], message.as_string())


class EmailSender(Protocol):
    def send(self, payload: EmailDispatchPayload) -> None:
        ...


class OutlookEmailSender:
    def __init__(self, timeout: int = 15) -> None:
        try:
            import win32com.client  # type: ignore
        except ImportError as exc:
            raise RuntimeError("pywin32 não está disponível no ambiente.") from exc
        self._win32 = win32com.client
        self.timeout = timeout

    def send(self, payload: EmailDispatchPayload) -> None:
        import threading

        result: dict = {}

        def _do_send() -> None:
            try:
                import pythoncom
                pythoncom.CoInitialize()
                outlook = self._win32.Dispatch("Outlook.Application")
                mail_item = outlook.CreateItem(0)
                mail_item.To = payload.recipient
                mail_item.Subject = payload.subject
                mail_item.HTMLBody = payload.html_body
                mail_item.Send()
                result["ok"] = True
            except Exception as exc:
                result["error"] = exc
            finally:
                try:
                    import pythoncom
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

        thread = threading.Thread(target=_do_send, daemon=True)
        thread.start()
        thread.join(timeout=self.timeout)

        if thread.is_alive():
            raise RuntimeError(
                f"Outlook não respondeu em {self.timeout}s. "
                "Verifique se o Outlook está aberto e sem janelas de diálogo "
                "bloqueando (login, erro de perfil ou atualização pendente)."
            )
        if "error" in result:
            raise result["error"]


class AutoEmailSender:
    def __init__(self, smtp_config: Optional[SmtpServerConfig] = None) -> None:
        self.smtp_config = smtp_config

    def send(self, payload: EmailDispatchPayload) -> None:
        last_error: Optional[Exception] = None
        try:
            OutlookEmailSender().send(payload)
            return
        except Exception as exc:
            last_error = exc

        if self.smtp_config:
            SmtpEmailSender(self.smtp_config).send(payload)
            return

        if last_error:
            raise RuntimeError(
                "Outlook não está disponível e a configuração SMTP não foi preenchida."
            ) from last_error
        raise RuntimeError("Nenhum meio de envio disponível.")


class EmailDispatchController:
    def __init__(
        self,
        data_source: SpreadsheetDataSource,
        template_catalog: TemplateCatalog,
        merge_engine: MergeFieldEngine,
        html_builder: DynamicEmailHtmlBuilder,
        audit_user: Optional[str] = None,
    ) -> None:
        self.data_source = data_source
        self.template_catalog = template_catalog
        self.merge_engine = merge_engine
        self.html_builder = html_builder
        self.audit_user = audit_user  # Para logging LGPD

    def load_source(self, file_path: str) -> list[str]:
        return self.data_source.load_file(file_path)

    def select_sheet(self, sheet_name: str) -> None:
        self.data_source.select_sheet(sheet_name)

    def list_columns(self) -> list[str]:
        return self.data_source.list_columns()

    def list_templates(self) -> list[EmailTemplateModel]:
        return self.template_catalog.list_templates()

    def get_template(self, template_name: str) -> EmailTemplateModel:
        return self.template_catalog.get_template(template_name)

    def merge_subject(self, subject: str, row_data: dict[str, str]) -> str:
        return self.merge_engine.render(subject, row_data)

    def merge_body(self, body: str, row_data: dict[str, str]) -> str:
        merged = self.merge_engine.render(body, row_data)
        return _sanitize_email_body_text(merged)

    def build_preview_html(self, subject: str, body: str, row_data: Optional[dict[str, str]] = None) -> str:
        return self.html_builder.build_html(
            subject=subject,
            plain_body=_sanitize_email_body_text(body),
            custom_data=row_data,
        )

    def build_payload(self, recipient: str, subject: str, body: str, row_data: dict[str, str]) -> EmailDispatchPayload:
        merged_subject = self.merge_subject(subject, row_data)
        merged_body = self.merge_body(body, row_data)
        html_body = self.build_preview_html(merged_subject, merged_body, row_data)
        return EmailDispatchPayload(recipient=recipient, subject=merged_subject, html_body=html_body)

    def dispatch_batch(
        self,
        sender: EmailSender,
        recipient_column: str,
        subject: str,
        body: str,
    ) -> tuple[int, int]:  # Retorna (sent_count, failed_count)
        """
        Envia um lote de emails com logging LGPD integrado.
        
        Returns:
            Tupla (sent_count, failed_count)
        """
        sent_count = 0
        failed_count = 0
        
        for row in self.data_source.iter_rows():
            recipient = row.get(recipient_column, "").strip()
            if not recipient:
                continue
            
            try:
                sender.send(self.build_payload(recipient, subject, body, row))
                sent_count += 1
                
                # Registrar envio bem-sucedido em auditoria LGPD
                if HAS_AUDIT_LOGGER and audit_logger and self.audit_user:
                    # Extrair matrícula, cidade e faturas do row_data
                    matricula = self._extract_value(row, ["Matrícula", "matricula", "MATRICULA", "Matrícula"])
                    city = self._extract_city(row)
                    open_invoices = self._extract_invoices(row)
                    
                    audit_logger.log_email_dispatch(
                        user=self.audit_user,
                        recipient_email=recipient,
                        matricula=matricula,
                        open_invoices=open_invoices,
                        email_subject=subject,
                        city=city,
                        success=True,
                    )
            
            except Exception as send_exc:
                failed_count += 1
                
                # Registrar erro em auditoria LGPD
                if HAS_AUDIT_LOGGER and audit_logger and self.audit_user:
                    matricula = self._extract_value(row, ["Matrícula", "matricula", "MATRICULA", "Matrícula"])
                    city = self._extract_city(row)
                    open_invoices = self._extract_invoices(row)
                    
                    audit_logger.log_email_dispatch(
                        user=self.audit_user,
                        recipient_email=recipient,
                        matricula=matricula,
                        open_invoices=open_invoices,
                        email_subject=subject,
                        city=city,
                        success=False,
                        error_detail=str(send_exc)
                    )
        
        return sent_count, failed_count
    
    def _extract_value(self, row: dict[str, str], possible_keys: list[str]) -> str:
        """Extrai um valor do row para diferentes variações de chave."""
        for key in possible_keys:
            value = row.get(key, "").strip()
            if value:
                return value
        return ""
    
    def _extract_city(self, row: dict[str, str]) -> str:
        """
        Extrai cidade do row_data.
        
        Tenta múltiplas variações de nomes de coluna para cidade.
        """
        city_patterns = [
            "Cidade", "cidade", "CIDADE", "City",
            "Localidade", "localidade", "Município", "município"
        ]
        return self._extract_value(row, city_patterns)
    
    def _extract_invoices(self, row: dict[str, str]) -> list[str]:
        """
        Extrai lista de faturas em aberto do row_data.
        
        Tenta múltiplas variações de nomes de coluna para faturas.
        Suporta:
        - Campo único com valores separados por vírgula ou ponto-e-vírgula
        - Múltiplos campos numerados (Fatura1, Fatura2, etc)
        - IDs numéricos de fatura (84635086, 52680450, etc)
        """
        invoices = []
        
        # Variações de nomes que procuramos (incluindo ID_FATURA)
        invoice_patterns = [
            "Faturas", "faturas", "FATURAS", 
            "Fatura_em_aberto", "Fatura em aberto", "fatura_em_aberto",
            "ID_FATURA", "Id_Fatura", "id_fatura", "ID_Faturas",
            "invoices", "invoices_open",
            "NumerosFaturas", "numeros_faturas", "Números_Faturas"
        ]
        
        # Procurar por campo consolidado
        for pattern in invoice_patterns:
            value = row.get(pattern, "").strip()
            if value:
                # Dividir por vírgula ou ponto-e-vírgula
                parts = re.split(r'[,;]', value)
                invoices.extend([p.strip() for p in parts if p.strip()])
                if invoices:
                    return invoices
        
        # Procurar por campos numerados (Fatura1, Fatura2, etc)
        i = 1
        while i <= 20:  # Verificar até 20 faturas
            for key in row.keys():
                if re.match(rf'Fatura\s*{i}|fatura\s*{i}', key, re.IGNORECASE):
                    value = row[key].strip()
                    if value and value.lower() not in ['', 'nan', 'none', '0', '-']:
                        invoices.append(value)
            i += 1
        
        return invoices


@dataclass(slots=True)
class EmailDispatchTheme:
    bg: str = "#EDF2F7"
    panel: str = "#FFFFFF"
    panel_alt: str = "#F7FAFC"
    border: str = "#D9E2EC"
    text: str = "#102A43"
    muted: str = "#627D98"
    accent: str = "#F27D16"
    accent_soft: str = "#FFF3E8"
    success: str = "#1F9D55"


@dataclass(slots=True)
class EmailDispatchModuleConfig:
    title: str = "Disparo de E-mails"
    brand_name: str = "Mathools"
    recipient_column_candidates: list[str] = field(
        default_factory=lambda: ["Email", "EMAIL", "E-mail", "email"]
    )
    templates: Optional[list[EmailTemplateModel]] = None
    enable_docx_support: bool = True
    default_delivery_mode: str = "auto"


class EmailDispatchModule:
    def __init__(
        self,
        page: ft.Page,
        config: Optional[EmailDispatchModuleConfig] = None,
        on_back: Optional[Callable[[ft.ControlEvent], None]] = None,
        ic_func: Optional[Callable[[str], str]] = None,
        audit_user: Optional[str] = None,
    ) -> None:
        self.page = page
        self.config = config or EmailDispatchModuleConfig()
        self.theme = EmailDispatchTheme()
        self.audit_user = audit_user  # Para logging LGPD
        self.controller = EmailDispatchController(
            data_source=SpreadsheetDataSource(),
            template_catalog=TemplateCatalog(self.config.templates),
            merge_engine=MergeFieldEngine(),
            html_builder=DynamicEmailHtmlBuilder(),
            audit_user=self.audit_user,
        )
        self.on_back = on_back
        self.ic = ic_func or (lambda name: getattr(ft.Icons, name, name.lower()))

        self.docx_engine: Optional[DocxMergeEngine] = None
        if HAS_DOCX_MERGE and self.config.enable_docx_support:
            self.docx_engine = DocxMergeEngine()

        self.docx_file_label = ft.Text("Nenhum documento DOCX carregado", color=self.theme.muted, size=12)
        self.docx_fields_display = ft.Text("", color=self.theme.accent, size=11)
        self.docx_preview_html = _flet_html_control(
            "<p style='color: #627D98; font-size: 13px;'>Carregue um documento DOCX para visualizar.</p>"
        )

        self.file_label = ft.Text("Nenhum arquivo selecionado", color=self.theme.muted, size=12)
        self.companies = DEFAULT_COMPANIES
        self.city_dropdown = ft.Dropdown(
            label="Unidade / Empresa",
            options=[ft.dropdown.Option(key=k, text=v.name) for k, v in self.companies.items()],
            value="Itapoá",
            dense=True,
            border_radius=14,
            border_color=self.theme.border,
            bgcolor=self.theme.panel,
            on_change=self._on_city_change,
            text_style=ft.TextStyle(color=self.theme.text),
        )
        self.sheet_dropdown = ft.Dropdown(
            label="Origem / Aba",
            options=[],
            dense=True,
            border_radius=14,
            border_color=self.theme.border,
            bgcolor=self.theme.panel,
            on_change=self._on_sheet_change,
            text_style=ft.TextStyle(color=self.theme.text),
        )
        self.template_dropdown = ft.Dropdown(
            label="Modelo rápido",
            dense=True,
            border_radius=14,
            border_color=self.theme.border,
            bgcolor=self.theme.panel,
            on_change=self._on_template_change,
            text_style=ft.TextStyle(color=self.theme.text),
        )
        self.subject_field = ft.TextField(
            label="Assunto",
            border_radius=16,
            multiline=False,
            border_color=self.theme.border,
            bgcolor=self.theme.panel,
            color=self.theme.text,
            text_style=ft.TextStyle(color=self.theme.text),
        )
        self.body_field = ft.TextField(
            label="Corpo do e-mail",
            multiline=True,
            min_lines=18,
            max_lines=24,
            border_radius=18,
            border_color=self.theme.border,
            bgcolor=self.theme.panel,
            on_change=self._update_preview,
            color=self.theme.text,
            text_style=ft.TextStyle(color=self.theme.text),
        )
        self.recipient_dropdown = ft.Dropdown(
            label="Coluna de destinatário",
            dense=True,
            border_radius=14,
            border_color=self.theme.border,
            bgcolor=self.theme.panel,
            text_style=ft.TextStyle(color=self.theme.text),
        )
        self.delivery_mode_dropdown = ft.Dropdown(
            label="Modo de envio",
            value=self.config.default_delivery_mode,
            dense=True,
            border_radius=14,
            border_color=self.theme.border,
            bgcolor=self.theme.panel,
            options=[
                ft.dropdown.Option(key="auto", text="Automático (Outlook e SMTP)"),
                ft.dropdown.Option(key="outlook", text="Somente Outlook"),
                ft.dropdown.Option(key="smtp", text="Somente SMTP"),
            ],
            on_change=self._on_delivery_mode_change,
            text_style=ft.TextStyle(color=self.theme.text),
        )
        self.smtp_host_field = ft.TextField(label="SMTP Host", border_radius=16, border_color=self.theme.border, bgcolor=self.theme.panel, color=self.theme.text, text_style=ft.TextStyle(color=self.theme.text))
        self.smtp_port_field = ft.TextField(label="Porta SMTP", value="587", border_radius=16, border_color=self.theme.border, bgcolor=self.theme.panel, keyboard_type=ft.KeyboardType.NUMBER, color=self.theme.text, text_style=ft.TextStyle(color=self.theme.text))
        self.smtp_username_field = ft.TextField(label="Usuário SMTP", border_radius=16, border_color=self.theme.border, bgcolor=self.theme.panel, color=self.theme.text, text_style=ft.TextStyle(color=self.theme.text))
        self.smtp_password_field = ft.TextField(label="Senha SMTP", password=True, can_reveal_password=True, border_radius=16, border_color=self.theme.border, bgcolor=self.theme.panel, color=self.theme.text, text_style=ft.TextStyle(color=self.theme.text))
        self.smtp_from_name_field = ft.TextField(label="Nome do remetente", value=self.config.brand_name, border_radius=16, border_color=self.theme.border, bgcolor=self.theme.panel, color=self.theme.text, text_style=ft.TextStyle(color=self.theme.text))
        self.smtp_from_email_field = ft.TextField(label="Email do remetente", border_radius=16, border_color=self.theme.border, bgcolor=self.theme.panel, color=self.theme.text, text_style=ft.TextStyle(color=self.theme.text))
        self.smtp_tls_checkbox = ft.Checkbox(label="Usar STARTTLS", value=True)
        self.smtp_settings_panel = ft.Container()
        self.template_hint = ft.Text("", color=self.theme.muted, size=12)
        self.preview_frame = ft.Container(
            bgcolor=self.theme.panel,
            border=ft.Border.all(1, self.theme.border),
            border_radius=20,
            padding=24,
            content=ft.Text("Carregue um arquivo para gerar a prévia.", color=self.theme.muted),
        )
        self._preview_table_data = ft.DataTable(
            columns=[ft.DataColumn(ft.Text(""))],
            rows=[],
            heading_row_color=self.theme.panel_alt,
            data_row_min_height=42,
            data_row_max_height=52,
            visible=False,
        )
        self._preview_table_empty = ft.Text("Carregue um arquivo para ver os dados.", color=self.theme.muted, size=12)
        self.preview_table = ft.Column(controls=[self._preview_table_empty, self._preview_table_data], scroll=ft.ScrollMode.AUTO)
        self.fields_wrap = ft.Row(wrap=True, spacing=8, run_spacing=8)
        self.status_text = ft.Text("", color=self.theme.muted, size=12)

        # Campos para controlar comportamento de download de faturas (UI help)
        self.batch_size_field = ft.TextField(
            label="Faturas baixadas ao mesmo tempo",
            value="10",
            border_radius=12,
            keyboard_type=ft.KeyboardType.NUMBER,
            width=220,
            text_style=ft.TextStyle(color=self.theme.text),
        )
        self.prepared_count_field = ft.TextField(
            label="Faturas preparadas antecipadamente",
            value="15",
            border_radius=12,
            keyboard_type=ft.KeyboardType.NUMBER,
            width=220,
            text_style=ft.TextStyle(color=self.theme.text),
        )

        self._populate_template_dropdown()
        # Não forçar empresa padrão — a empresa é determinada por CIDADE_IMOVEL em cada linha
        self._refresh_delivery_mode_ui()

    def build(self) -> ft.Control:
        return ft.Container(
            expand=True,
            bgcolor=self.theme.bg,
            padding=24,
            content=ft.Column(
                expand=True,
                spacing=24,
                controls=[
                    self._build_header(),
                    self._build_workflow_steps(),
                    ft.ResponsiveRow(
                        controls=[
                            ft.Container(col={"xs": 12, "lg": 7}, content=self._build_main_panel()),
                            ft.Container(col={"xs": 12, "lg": 5}, content=self._build_sidebar()),
                        ]
                    ),
                ],
            ),
        )

    def _build_header(self) -> ft.Control:
        actions = []
        if self.on_back:
            actions.append(
                ft.TextButton(
                    content=ft.Row(controls=[ft.Icon(self.ic("ARROW_BACK_ROUNDED"), size=18), ft.Text("Voltar")], tight=True, spacing=6),
                    style=ft.ButtonStyle(color=self.theme.text),
                    on_click=self.on_back,
                )
            )
        actions.append(
            ft.OutlinedButton(
                content=ft.Row(controls=[ft.Icon(self.ic("AUTO_AWESOME_ROUNDED"), size=18), ft.Text("Atualizar prévia")], tight=True, spacing=6),
                style=ft.ButtonStyle(color=self.theme.text, side=ft.BorderSide(1, self.theme.border), shape=ft.RoundedRectangleBorder(radius=16)),
                on_click=self._update_preview,
            )
        )
        return ft.Container(
            bgcolor=self.theme.panel,
            border_radius=24,
            border=ft.Border.all(1, self.theme.border),
            padding=24,
            content=ft.Row(
                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                vertical_alignment=ft.CrossAxisAlignment.CENTER,
                controls=[
                    ft.Column(spacing=6, controls=[
                        ft.Text(self.config.title, size=28, weight=ft.FontWeight.W_700, color=self.theme.text),
                        ft.Text("Editor de mala direta com leitura segura de planilhas e HTML corporativo.", size=13, color=self.theme.muted),
                        self.file_label,
                    ]),
                    ft.Row(spacing=10, controls=actions),
                ],
            ),
        )

    def _on_workflow_step_click(self, step: int) -> None:
        if step == 1:
            self.status_text.value = "👉 Clique em 'Selecionar Arquivo' para carregar sua planilha"
        elif step == 2:
            self.status_text.value = "👉 Escolha o modelo e a coluna de destinatários" if self.controller.data_source.file_path else "⚠️ Carregue um arquivo primeiro"
        elif step == 3:
            self.status_text.value = "👉 Edite o assunto e o corpo do email" if (self.template_dropdown.value and self.recipient_dropdown.value) else "⚠️ Configure os campos anteriores primeiro"
        elif step == 4:
            self.status_text.value = "👉 Revise e clique em 'Enviar Emails'" if ((self.subject_field.value or "").strip() and (self.body_field.value or "").strip()) else "⚠️ Personalize o email primeiro"
        try:
            self.status_text.update()
        except Exception:
            pass

    def _build_workflow_steps(self) -> ft.Control:
        steps = [
            ("📁 Carregar", "Selecione sua planilha", self._is_step_complete(1)),
            ("📋 Configurar", "Escolha template e destinatários", self._is_step_complete(2)),
            ("✏️ Personalizar", "Edite assunto e mensagem", self._is_step_complete(3)),
            ("🚀 Enviar", "Dispare os emails", self._is_step_complete(4)),
        ]
        step_controls = []
        for i, (icon, desc, complete) in enumerate(steps, 1):
            color = self.theme.success if complete else (self.theme.accent if i == self._get_current_step() else self.theme.muted)
            step_controls.append(
                ft.Container(
                    bgcolor=self.theme.panel,
                    border_radius=12,
                    border=ft.Border.all(2, color),
                    padding=ft.Padding.symmetric(horizontal=16, vertical=12),
                    content=ft.Column(
                        spacing=4,
                        horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                        controls=[
                            ft.Text(icon, size=20),
                            ft.Text(f"Passo {i}", size=12, weight=ft.FontWeight.W_600, color=color),
                            ft.Text(desc, size=11, color=self.theme.muted, text_align=ft.TextAlign.CENTER),
                        ],
                    ),
                    ink=True,
                    on_click=lambda e, step_num=i: self._on_workflow_step_click(step_num),
                )
            )
        return ft.Container(
            bgcolor=self.theme.panel,
            border_radius=16,
            border=ft.Border.all(1, self.theme.border),
            padding=20,
            content=ft.Column(spacing=12, controls=[
                ft.Text("📋 Fluxo de Trabalho", size=16, weight=ft.FontWeight.W_700, color=self.theme.text),
                ft.Row(spacing=12, controls=step_controls, alignment=ft.MainAxisAlignment.CENTER),
            ]),
        )

    def _is_step_complete(self, step: int) -> bool:
        if step == 1:
            return self.controller.data_source.file_path is not None
        elif step == 2:
            return self.template_dropdown.value is not None and self.recipient_dropdown.value is not None
        elif step == 3:
            return bool((self.subject_field.value or "").strip() and (self.body_field.value or "").strip())
        return False

    def _get_current_step(self) -> int:
        for step in (1, 2, 3):
            if not self._is_step_complete(step):
                return step
        return 4

    def _build_main_panel(self) -> ft.Control:
        controls = [self._build_file_section()]
        if self.config.enable_docx_support and HAS_DOCX_MERGE:
            controls.append(self._build_docx_section())
        controls.extend([self._build_config_section(), self._build_compose_section(), self._build_actions_section()])
        return ft.Container(bgcolor=self.theme.panel, border_radius=20, border=ft.Border.all(1, self.theme.border), padding=24, content=ft.Column(spacing=24, controls=controls))

    def _build_docx_section(self) -> ft.Control:
        if not self.docx_engine:
            return ft.Container()
        is_loaded = self.docx_engine.data_source.template and self.docx_engine.data_source.template.is_loaded
        docx_status = "✅ Documento carregado" if is_loaded else "❌ Nenhum documento carregado"
        return ft.Container(
            bgcolor=self.theme.panel_alt, border_radius=16, border=ft.Border.all(1, self.theme.border), padding=20,
            content=ft.Column(spacing=12, controls=[
                ft.Row(controls=[ft.Icon(ft.Icons.DESCRIPTION, color=self.theme.accent, size=24), ft.Text("1.5. Carregar Documento DOCX", size=18, weight=ft.FontWeight.W_700, color=self.theme.text)], spacing=12, vertical_alignment=ft.CrossAxisAlignment.CENTER),
                ft.Text("Selecione um documento Word (.docx) para usar como modelo no email.", size=13, color=self.theme.muted),
                ft.Container(bgcolor=self.theme.bg, border_radius=12, padding=16, content=ft.Row(
                    controls=[
                        ft.Icon(ft.Icons.CHECK_CIRCLE if is_loaded else ft.Icons.ERROR, color=self.theme.success if is_loaded else "#D64545", size=20),
                        ft.Text(docx_status, size=14, color=self.theme.text),
                        ft.Container(expand=True),
                        ft.ElevatedButton("Carregar DOCX", icon=ft.Icons.UPLOAD_FILE, style=ft.ButtonStyle(bgcolor=self.theme.accent, color=ft.Colors.WHITE, shape=ft.RoundedRectangleBorder(radius=12)), on_click=self._show_docx_picker_dialog),
                    ], spacing=12, vertical_alignment=ft.CrossAxisAlignment.CENTER,
                )),
                self.docx_file_label, self.docx_fields_display,
                ft.Text("Prévia do Documento DOCX", size=16, weight=ft.FontWeight.W_600, color=self.theme.text),
                ft.Container(height=300, bgcolor=self.theme.panel, border=ft.Border.all(1, self.theme.border), border_radius=12, padding=16, content=self.docx_preview_html),
            ]),
        )

    def _build_file_section(self) -> ft.Control:
        is_loaded = self.controller.data_source.file_path is not None
        file_status = "✅ Arquivo carregado" if is_loaded else "❌ Nenhum arquivo selecionado"
        return ft.Container(
            bgcolor=self.theme.panel_alt, border_radius=16, border=ft.Border.all(1, self.theme.border), padding=20,
            content=ft.Column(spacing=12, controls=[
                ft.Row(controls=[ft.Icon(ft.Icons.FOLDER_OPEN, color=self.theme.accent, size=24), ft.Text("1. Carregar Planilha", size=18, weight=ft.FontWeight.W_700, color=self.theme.text)], spacing=12, vertical_alignment=ft.CrossAxisAlignment.CENTER),
                ft.Text("Selecione um arquivo CSV ou Excel contendo os dados dos destinatários.", size=13, color=self.theme.muted),
                ft.Container(bgcolor=self.theme.bg, border_radius=12, padding=16, content=ft.Row(
                    controls=[
                        ft.Icon(ft.Icons.CHECK_CIRCLE if is_loaded else ft.Icons.ERROR, color=self.theme.success if is_loaded else "#D64545", size=20),
                        ft.Text(file_status, size=14, color=self.theme.text),
                        ft.Container(expand=True),
                        ft.ElevatedButton("Selecionar Arquivo", icon=ft.Icons.UPLOAD_FILE, style=ft.ButtonStyle(bgcolor=self.theme.accent, color=ft.Colors.WHITE, shape=ft.RoundedRectangleBorder(radius=12)), on_click=self._show_file_picker_dialog),
                    ], spacing=12, vertical_alignment=ft.CrossAxisAlignment.CENTER,
                )),
                self.sheet_dropdown,
            ]),
        )

    def _build_config_section(self) -> ft.Control:
        return ft.Container(
            bgcolor=self.theme.panel_alt, border_radius=16, border=ft.Border.all(1, self.theme.border), padding=20,
            content=ft.Column(spacing=12, controls=[
                ft.Row(controls=[ft.Icon(ft.Icons.SETTINGS, color=self.theme.accent, size=24), ft.Text("2. Configurar Envio", size=18, weight=ft.FontWeight.W_700, color=self.theme.text)], spacing=12, vertical_alignment=ft.CrossAxisAlignment.CENTER),
                ft.Text("Escolha a unidade, o template de email e a coluna que contém os endereços de email.", size=13, color=self.theme.muted),
                ft.ResponsiveRow(controls=[
                    ft.Container(col={"xs": 12, "md": 6}, content=self.city_dropdown),
                    ft.Container(col={"xs": 12, "md": 6}, content=self.template_dropdown),
                    ft.Container(col={"xs": 12, "md": 6}, content=self.recipient_dropdown),
                    ft.Container(col={"xs": 12, "md": 6}, content=self.delivery_mode_dropdown),
                ]),
                self.smtp_settings_panel,
                self.template_hint,
            ]),
        )

    def _build_compose_section(self) -> ft.Control:
        return ft.Container(
            bgcolor=self.theme.panel_alt, border_radius=16, border=ft.Border.all(1, self.theme.border), padding=20,
            content=ft.Column(spacing=16, controls=[
                ft.Row(controls=[ft.Icon(ft.Icons.EDIT, color=self.theme.accent, size=24), ft.Text("3. Personalizar Mensagem", size=18, weight=ft.FontWeight.W_700, color=self.theme.text)], spacing=12, vertical_alignment=ft.CrossAxisAlignment.CENTER),
                ft.Text("Edite o assunto e o corpo do email. Use as variáveis disponíveis para personalizar cada mensagem.", size=13, color=self.theme.muted),
                self.subject_field,
                self.body_field,
                ft.Text("Prévia do Email", size=16, weight=ft.FontWeight.W_600, color=self.theme.text),
                ft.Container(height=300, content=self.preview_frame),
            ]),
        )

    def _build_actions_section(self) -> ft.Control:
        can_send = self._is_step_complete(3)
        has_docx = self.docx_engine is not None and self.docx_engine.data_source.template and self.docx_engine.data_source.template.is_loaded
        return ft.Container(
            bgcolor=self.theme.panel_alt, border_radius=16, border=ft.Border.all(1, self.theme.border), padding=20,
            content=ft.Column(spacing=16, controls=[
                ft.Row(controls=[ft.Icon(ft.Icons.SEND, color=self.theme.accent, size=24), ft.Text("4. Enviar Emails", size=18, weight=ft.FontWeight.W_700, color=self.theme.text)], spacing=12, vertical_alignment=ft.CrossAxisAlignment.CENTER),
                ft.Text("Revise tudo e clique em 'Enviar Emails' para iniciar o disparo.", size=13, color=self.theme.muted),
                ft.Container(bgcolor=self.theme.bg, border_radius=12, padding=16, content=ft.Column(spacing=12, controls=[
                    ft.Text("📎 Anexar Documento ao Email", size=14, weight=ft.FontWeight.W_600, color=self.theme.text),
                    ft.Row(controls=[
                        ft.Checkbox(label="Anexar documento DOCX renderizado", value=has_docx, on_change=self._on_attach_docx_change),
                        ft.Container(expand=True),
                        ft.Text(f"Documento: {self.docx_engine.data_source.template.file_path.name if has_docx else 'Nenhum'}", size=12, color=self.theme.muted),
                    ], vertical_alignment=ft.CrossAxisAlignment.CENTER),
                ])),
                ft.Container(bgcolor=self.theme.bg, border_radius=12, padding=16, content=ft.Row(
                    controls=[
                        ft.ElevatedButton("📧 Enviar Emails", icon=ft.Icons.SEND, disabled=not can_send, style=ft.ButtonStyle(bgcolor=self.theme.success if can_send else self.theme.muted, color=ft.Colors.WHITE, shape=ft.RoundedRectangleBorder(radius=12)), on_click=self._send_emails),
                        ft.Container(expand=True),
                        ft.Text(f"Pronto para enviar para {self._get_recipient_count()} destinatários" if can_send else "Complete os passos anteriores", size=13, color=self.theme.success if can_send else self.theme.muted),
                    ], spacing=12, vertical_alignment=ft.CrossAxisAlignment.CENTER,
                )),
                self.status_text,
            ]),
        )

    def _build_sidebar(self) -> ft.Control:
        return ft.Container(
            bgcolor=self.theme.panel, border_radius=20, border=ft.Border.all(1, self.theme.border), padding=20,
            content=ft.Column(spacing=20, controls=[
                ft.Container(bgcolor=self.theme.panel_alt, border_radius=16, border=ft.Border.all(1, self.theme.border), padding=16, content=ft.Column(spacing=12, controls=[
                    ft.Row(controls=[ft.Icon(ft.Icons.LABEL, color=self.theme.accent, size=20), ft.Text("Variáveis Disponíveis", size=16, weight=ft.FontWeight.W_600, color=self.theme.text)], spacing=8, vertical_alignment=ft.CrossAxisAlignment.CENTER),
                    ft.Text("Clique em uma coluna para inserir {{variável}} no texto do email.", size=12, color=self.theme.muted),
                    ft.Container(content=self.fields_wrap, height=120, alignment=ft.Alignment(-1, -1)),
                ])),
                ft.Container(bgcolor=self.theme.panel_alt, border_radius=16, border=ft.Border.all(1, self.theme.border), padding=16, content=ft.Column(spacing=12, controls=[
                    ft.Row(controls=[ft.Icon(ft.Icons.TABLE_CHART, color=self.theme.accent, size=20), ft.Text("Prévia dos Dados", size=16, weight=ft.FontWeight.W_600, color=self.theme.text)], spacing=8, vertical_alignment=ft.CrossAxisAlignment.CENTER),
                    ft.Text("Confira os primeiros registros da sua planilha.", size=12, color=self.theme.muted),
                    ft.Container(content=self.preview_table, height=250),
                ])),
                # Controles de ajuste de download de faturas com ajuda (tooltip)
                ft.Container(bgcolor=self.theme.panel_alt, border_radius=16, border=ft.Border.all(1, self.theme.border), padding=16, content=ft.Column(spacing=12, controls=[
                    ft.Row(controls=[ft.Icon(ft.Icons.DOWNLOAD, color=self.theme.accent, size=20), ft.Text("Config. de Download de Faturas", size=16, weight=ft.FontWeight.W_600, color=self.theme.text)], spacing=8, vertical_alignment=ft.CrossAxisAlignment.CENTER),
                    ft.Text("Ajuste como as faturas serão preparadas e baixadas pelo sistema.", size=12, color=self.theme.muted),
                    ft.Row(controls=[
                        ft.Container(content=self.batch_size_field),
                        ft.IconButton(ft.Icons.INFO_OUTLINE, tooltip=(
                            "Quantas faturas o sistema tenta baixar ao mesmo tempo. "
                            "Um valor maior pode acelerar, mas também aumenta o risco de bloqueio pelo Waterfy."
                        )),
                    ], alignment=ft.MainAxisAlignment.START, spacing=8),
                    ft.Row(controls=[
                        ft.Container(content=self.prepared_count_field),
                        ft.IconButton(ft.Icons.INFO_OUTLINE, tooltip=(
                            "Quantas faturas o sistema deixa prontas antes de começar a baixar. "
                            "Isso forma uma fila de preparação; valores maiores diminuem intervalos, mas podem deixar o servidor sobrecarregado." 
                        )),
                    ], alignment=ft.MainAxisAlignment.START, spacing=8),
                    ft.Text("Cuidado: aumentar demais pode provocar bloqueios no Waterfy.", size=12, color="#D97706"),
                ])),
            ]),
        )

    def _populate_template_dropdown(self) -> None:
        templates = self.controller.list_templates()
        self.template_dropdown.options = [ft.dropdown.Option(key=t.name, text=t.name) for t in templates]
        if templates:
            self.template_dropdown.value = templates[0].name
            self._apply_template(templates[0])

    def _apply_template(self, template: EmailTemplateModel) -> None:
        self.subject_field.value = template.subject
        self.body_field.value = _sanitize_email_body_text(template.body)
        self.template_hint.value = template.description

    def _on_city_change(self, event: ft.ControlEvent) -> None:
        """
        O seletor de unidade afeta apenas a prévia visual da UI.
        A empresa real de cada email é determinada pela coluna CIDADE_IMOVEL
        na planilha, resolvida pelo email_merge_module.
        """
        if not self.city_dropdown.value:
            return
        selected_company = self.companies.get(self.city_dropdown.value)
        if selected_company:
            self.controller.html_builder.set_company_config(selected_company)
        self._update_preview()

    def _on_file_picked(self, event: ft.FilePickerResultEvent) -> None:
        if not event.files:
            return
        file_path = event.files[0].path
        try:
            sheet_names = self.controller.load_source(file_path)
            self.file_label.value = Path(file_path).name
            self.sheet_dropdown.options = [ft.dropdown.Option(key=s, text=s) for s in sheet_names]
            if sheet_names:
                self.sheet_dropdown.value = sheet_names[0]
                self.controller.select_sheet(sheet_names[0])
            self._refresh_columns()
            self._refresh_preview_table()
            self.status_text.value = "Origem carregada com sucesso."
            self.status_text.color = self.theme.success
        except Exception as exc:
            self.status_text.value = f"Falha ao carregar arquivo: {exc}"
            self.status_text.color = "#D64545"
        self._update_preview()

    def _on_sheet_change(self, event: ft.ControlEvent) -> None:
        selected_sheet = event.data if event and event.data else self.sheet_dropdown.value
        if not selected_sheet:
            return
        self.sheet_dropdown.value = selected_sheet
        self.controller.select_sheet(selected_sheet)
        self._refresh_columns()
        self._refresh_preview_table()
        self._update_preview()

    def _on_template_change(self, event: ft.ControlEvent) -> None:
        if not self.template_dropdown.value:
            return
        self._apply_template(self.controller.get_template(self.template_dropdown.value))
        self._update_preview()

    def _on_delivery_mode_change(self, event: ft.ControlEvent) -> None:
        self._refresh_delivery_mode_ui()
        self.page.update()

    def _refresh_delivery_mode_ui(self) -> None:
        smtp_visible = self.delivery_mode_dropdown.value in {"auto", "smtp"}
        self.smtp_settings_panel.content = ft.Column(
            visible=smtp_visible, spacing=12,
            controls=[
                ft.Text("Configuração SMTP", size=14, weight=ft.FontWeight.W_600, color=self.theme.text),
                ft.Text("No modo automático, o sistema tenta Outlook primeiro e usa SMTP se a configuração estiver preenchida.", size=12, color=self.theme.muted),
                ft.ResponsiveRow(controls=[
                    ft.Container(col={"xs": 12, "md": 8}, content=self.smtp_host_field),
                    ft.Container(col={"xs": 12, "md": 4}, content=self.smtp_port_field),
                    ft.Container(col={"xs": 12, "md": 6}, content=self.smtp_username_field),
                    ft.Container(col={"xs": 12, "md": 6}, content=self.smtp_password_field),
                    ft.Container(col={"xs": 12, "md": 6}, content=self.smtp_from_name_field),
                    ft.Container(col={"xs": 12, "md": 6}, content=self.smtp_from_email_field),
                    ft.Container(col={"xs": 12, "md": 12}, content=self.smtp_tls_checkbox),
                ]),
            ],
        )

    def _build_smtp_config(self) -> Optional[SmtpServerConfig]:
        host = (self.smtp_host_field.value or "").strip()
        port_raw = (self.smtp_port_field.value or "").strip()
        username = (self.smtp_username_field.value or "").strip()
        password = self.smtp_password_field.value or ""
        from_name = (self.smtp_from_name_field.value or "").strip() or self.config.brand_name
        from_email = (self.smtp_from_email_field.value or "").strip() or None

        if not any([host, port_raw, username, password, from_email]):
            return None
        if not all([host, port_raw, username, password]):
            raise ValueError("Preencha host, porta, usuário e senha do SMTP.")
        try:
            port = int(port_raw)
        except ValueError as exc:
            raise ValueError("A porta SMTP deve ser numérica.") from exc

        return SmtpServerConfig(host=host, port=port, username=username, password=password, use_tls=bool(self.smtp_tls_checkbox.value), from_name=from_name, from_email=from_email)

    def _build_sender(self) -> EmailSender:
        mode = self.delivery_mode_dropdown.value or "auto"
        smtp_config = self._build_smtp_config()
        if mode == "outlook":
            return OutlookEmailSender()
        if mode == "smtp":
            if not smtp_config:
                raise ValueError("Preencha a configuração SMTP para usar este modo.")
            return SmtpEmailSender(smtp_config)
        return AutoEmailSender(smtp_config=smtp_config)

    def _refresh_columns(self) -> None:
        columns = self.controller.list_columns()
        self.fields_wrap.controls = [
            ft.OutlinedButton(
                content=ft.Text(col),
                style=ft.ButtonStyle(color=self.theme.text, side=ft.BorderSide(1, self.theme.border), bgcolor=self.theme.accent_soft, shape=ft.RoundedRectangleBorder(radius=999)),
                on_click=lambda _, c=col: self._insert_variable(c),
            )
            for col in columns
        ]
        self.recipient_dropdown.options = [ft.dropdown.Option(key=col, text=col) for col in columns]
        self.recipient_dropdown.value = self._guess_recipient_column(columns)
        self.page.update()

    def _refresh_preview_table(self) -> None:
        rows = self.controller.data_source.preview_rows()
        if not rows:
            self._preview_table_data.visible = False
            self._preview_table_empty.visible = True
            self.page.update()
            return
        column_names = list(rows[0].keys())[:5]
        self._preview_table_data.columns = [ft.DataColumn(ft.Text(c, color=self.theme.text, weight=ft.FontWeight.W_600)) for c in column_names]
        self._preview_table_data.rows = [ft.DataRow(cells=[ft.DataCell(ft.Text(row.get(c, ""), size=12, color=self.theme.text)) for c in column_names]) for row in rows]
        self._preview_table_data.visible = True
        self._preview_table_empty.visible = False
        self.page.update()

    def _insert_variable(self, column_name: str) -> None:
        current_text = self.body_field.value or ""
        cursor_pos = self.body_field.cursor_pos if hasattr(self.body_field, "cursor_pos") and self.body_field.cursor_pos is not None else len(current_text)
        if cursor_pos is None or cursor_pos < 0:
            cursor_pos = len(current_text)
        token = f"{{{{{column_name}}}}}"
        before = current_text[:cursor_pos]
        after = current_text[cursor_pos:]
        if before and not before.endswith((" ", "\n", "\t")):
            before += " "
        self.body_field.value = _sanitize_email_body_text(before + token + after)
        self._update_preview()
        self.body_field.focus()

    def _show_docx_picker_dialog(self, event: ft.ControlEvent) -> None:
        try:
            if hasattr(ft, "FilePicker"):
                file_picker = ft.FilePicker(on_result=self._on_docx_picked)
                self.page.overlay.append(file_picker)
                file_picker.pick_files(dialog_title="Selecione um documento DOCX", allowed_extensions=["docx"])
            else:
                self.docx_file_label.value = "⚠️ Use Ctrl+O ou arraste um arquivo DOCX aqui"
                self.docx_file_label.color = "#F27D16"
                self.page.update()
        except Exception as exc:
            self.docx_file_label.value = f"Erro ao abrir seletor: {exc}"
            self.docx_file_label.color = "#D64545"
            self.page.update()

    def _on_docx_picked(self, event: ft.FilePickerResultEvent) -> None:
        if not event.files:
            return
        file_path = event.files[0].path
        try:
            if not self.docx_engine:
                self.docx_file_label.value = "Suporte a DOCX não disponível"
                self.docx_file_label.color = "#D64545"
                self.page.update()
                return
            success, message = self.docx_engine.load_template(file_path)
            if success:
                self.docx_file_label.value = Path(file_path).name
                self.docx_file_label.color = self.theme.success
                doc_info = self.docx_engine.get_document_info()
                fields = doc_info.get("fields", [])
                self.docx_fields_display.value = f"Campos encontrados: {', '.join(fields)}" if fields else "Nenhum campo {{...}} encontrado no documento"
                preview_html = self.docx_engine.get_preview_html()
                _set_flet_html_data(self.docx_preview_html, f"<div style='font-family:Arial,sans-serif;font-size:13px;line-height:1.6;color:#102A43;'>{preview_html}</div>")
                if fields:
                    self._add_docx_fields_to_sidebar(fields)
                self.status_text.value = f"✅ {message}"
                self.status_text.color = self.theme.success
            else:
                self.docx_file_label.value = f"❌ Erro: {message}"
                self.docx_file_label.color = "#D64545"
                self.status_text.value = f"❌ Erro ao carregar DOCX: {message}"
                self.status_text.color = "#D64545"
            self.page.update()
        except Exception as exc:
            self.docx_file_label.value = f"Erro ao carregar: {exc}"
            self.docx_file_label.color = "#D64545"
            self.status_text.value = f"❌ Erro: {exc}"
            self.status_text.color = "#D64545"
            self.page.update()

    def _on_attach_docx_change(self, event: ft.ControlEvent) -> None:
        if event.data == "true":
            self.status_text.value = "📎 Documento DOCX será anexado aos emails"
            self.status_text.color = self.theme.success
        else:
            self.status_text.value = "📧 Emails serão enviados sem anexo DOCX"
            self.status_text.color = self.theme.muted
        self.page.update()

    def _add_docx_fields_to_sidebar(self, fields: list[str]) -> None:
        existing_columns = set(self.controller.list_columns())
        for field_name in fields:
            if field_name not in existing_columns:
                self.fields_wrap.controls.append(
                    ft.OutlinedButton(
                        content=ft.Text(f"DOCX: {field_name}"),
                        style=ft.ButtonStyle(color=self.theme.accent, side=ft.BorderSide(1, self.theme.accent), bgcolor=self.theme.accent_soft, shape=ft.RoundedRectangleBorder(radius=999)),
                        on_click=lambda _, col=field_name: self._insert_variable(col),
                    )
                )
        self.page.update()

    def _show_file_picker_dialog(self, event: ft.ControlEvent) -> None:
        try:
            if hasattr(ft, "FilePicker"):
                file_picker = ft.FilePicker(on_result=self._on_file_picked)
                self.page.overlay.append(file_picker)
                file_picker.pick_files(dialog_title="Selecione um arquivo", allowed_extensions=["csv", "xlsx", "xls"])
            else:
                self.status_text.value = "⚠️ Para carregar arquivo, use Ctrl+O ou arraste um arquivo CSV/Excel aqui"
                self.status_text.color = "#F27D16"
                self.page.update()
        except Exception as exc:
            self.status_text.value = f"Erro ao abrir seletor: {exc}"
            self.status_text.color = "#D64545"
            self.page.update()

    def _get_recipient_count(self) -> int:
        try:
            return len(list(self.controller.data_source.iter_rows()))
        except Exception:
            return 0

    def _guess_recipient_column(self, columns: list[str]) -> Optional[str]:
        candidates = {c.lower(): c for c in self.config.recipient_column_candidates}
        for col in columns:
            if col.lower() in candidates:
                return col
        return columns[0] if columns else None

    def _get_preview_context(self) -> dict[str, str]:
        rows = self.controller.data_source.preview_rows(limit=1)
        return rows[0] if rows else {}

    def _update_preview(self, _: Optional[ft.ControlEvent] = None) -> None:
        try:
            context = self._get_preview_context()
            if not context:
                self.preview_frame.content = ft.Text("Carregue um arquivo para gerar a prévia personalizada.", color=self.theme.muted)
                self.page.update()
                return
            merged_subject = self.controller.merge_subject(self.subject_field.value or "", context)
            merged_body = self.controller.merge_body(self.body_field.value or "", context)
            html_content = self.controller.build_preview_html(merged_subject, merged_body, context)
            self.preview_frame.content = ft.Column(spacing=0, controls=[
                ft.Container(
                    padding=ft.Padding.symmetric(horizontal=16, vertical=10),
                    bgcolor="#F0F0F0",
                    border_radius=ft.BorderRadius.only(top_left=12, top_right=12),
                    content=ft.Row(controls=[
                        ft.Icon(ft.Icons.EMAIL, size=16, color="#666"),
                        ft.Text(merged_subject or "Assunto da mensagem", size=14, weight=ft.FontWeight.W_600, color="#333"),
                    ], spacing=8),
                ),
                _dispatch_email_body_preview_widget(html_content),
            ])
        except Exception as exc:
            self.preview_frame.content = ft.Text(f"Falha ao gerar prévia: {exc}", color="#D64545")
        self.page.update()

    def create_payloads(self) -> list[EmailDispatchPayload]:
        recipient_column = self.recipient_dropdown.value
        if not recipient_column:
            raise ValueError("Selecione a coluna de destinatário.")
        subject = self.subject_field.value or ""
        body = _sanitize_email_body_text(self.body_field.value or "")
        payloads: list[EmailDispatchPayload] = []
        for row in self.controller.data_source.iter_rows():
            recipient = row.get(recipient_column, "").strip()
            if not recipient or "@" not in recipient:
                continue
            payloads.append(self.controller.build_payload(recipient, subject, body, row))
        return payloads

    def _send_emails(self, event: ft.ControlEvent) -> None:
        """Processa e envia emails pelo provedor selecionado com logging LGPD integrado."""
        import time
        start_time = time.time()
        
        try:
            if not self.controller.data_source.file_path:
                self.status_text.value = "❌ Carregue um arquivo primeiro"
                self.status_text.color = "#D64545"
                self.page.update()
                return
            if not self.template_dropdown.value or not self.recipient_dropdown.value:
                self.status_text.value = "❌ Configure o template e selecione a coluna de destinatários"
                self.status_text.color = "#D64545"
                self.page.update()
                return
            if not (self.subject_field.value or "").strip() or not (self.body_field.value or "").strip():
                self.status_text.value = "❌ Preencha o assunto e o corpo do email"
                self.status_text.color = "#D64545"
                self.page.update()
                return

            recipient_column = self.recipient_dropdown.value
            subject = self.subject_field.value or ""
            body = _sanitize_email_body_text(self.body_field.value or "")
            
            # Contar total de destinatários válidos
            total_count = 0
            for row in self.controller.data_source.iter_rows():
                recipient = row.get(recipient_column, "").strip()
                if recipient and "@" in recipient:
                    total_count += 1
            
            if total_count == 0:
                self.status_text.value = "❌ Nenhum destinatário válido encontrado"
                self.status_text.color = "#D64545"
                self.page.update()
                return

            sender = self._build_sender()
            
            # Usar o novo dispatch_batch com logging integrado
            sent_count, failed_count = self.controller.dispatch_batch(
                sender=sender,
                recipient_column=recipient_column,
                subject=subject,
                body=body,
            )
            
            # Registrar lote em auditoria LGPD
            process_duration = time.time() - start_time
            if HAS_AUDIT_LOGGER and audit_logger and self.audit_user:
                audit_logger.log_batch_email_dispatch(
                    user=self.audit_user,
                    total_recipients=total_count,
                    successful_sends=sent_count,
                    failed_sends=failed_count,
                    process_duration_seconds=process_duration,
                    detail=f"Assunto: {subject[:50]}..."
                )
            
            # Atualizar status na interface
            if failed_count > 0:
                self.status_text.value = f"⚠️ {sent_count}/{total_count} email(s) enviados. Falhas: {failed_count}"
                self.status_text.color = "#F27D16"
            else:
                self.status_text.value = f"✅ {sent_count} email(s) enviados com sucesso."
                self.status_text.color = self.theme.success
            
            self.page.update()
        except ValueError as ve:
            self.status_text.value = f"❌ Erro na validação: {ve}"
            self.status_text.color = "#D64545"
            self.page.update()
        except Exception as exc:
            self.status_text.value = f"❌ Erro ao processar: {exc}"
            self.status_text.color = "#D64545"
            self.page.update()


def build_email_dispatch_module(
    page: ft.Page,
    on_back: Optional[Callable[[ft.ControlEvent], None]] = None,
    config: Optional[EmailDispatchModuleConfig] = None,
    ic_func: Optional[Callable[[str], str]] = None,
    audit_user: Optional[str] = None,
) -> EmailDispatchModule:
    """
    Constrói o módulo de disparo de emails com suporte a auditoria LGPD.
    
    Args:
        page: Página Flet
        on_back: Callback para botão voltar
        config: Configuração do módulo
        ic_func: Função para resolver ícones
        audit_user: Identificador do usuário para logging LGPD (ex: matrícula/login)
    
    Returns:
        Instância de EmailDispatchModule configurada
    """
    return EmailDispatchModule(
        page=page,
        on_back=on_back,
        config=config,
        ic_func=ic_func,
        audit_user=audit_user,
    )
