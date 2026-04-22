"""
Sanitizador de HTML
Remove conteúdo malicioso enquanto preserva formatação legítima
Conformidade com segurança (evita XSS em dashboards)
"""

import re
import json
from typing import List, Set, Dict, Any
from html.parser import HTMLParser
import logging
from config_loader import config

logger = logging.getLogger(__name__)


# Tags HTML permitidas (whitelist)
ALLOWED_TAGS = {
    # Estrutura
    'div', 'span', 'section', 'article', 'header', 'footer', 'nav',
    
    # Texto
    'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'strong', 'em', 'b', 'i',
    'u', 'br', 'hr', 'code', 'pre', 'blockquote', 'small',
    
    # Listas
    'ul', 'ol', 'li', 'dl', 'dt', 'dd',
    
    # Tabelas
    'table', 'thead', 'tbody', 'tfoot', 'tr', 'td', 'th', 'caption', 'colgroup', 'col',
    
    # Mídia (com cuidado)
    'img', 'svg', 'picture',
    
    # Formulários (read-only apenas)
    'label', 'fieldset', 'legend',
}

# Atributos permitidos por tag
ALLOWED_ATTRIBUTES: Dict[str, Set[str]] = {
    'a': {'href', 'title', 'target', 'rel'},
    'img': {'src', 'alt', 'width', 'height', 'loading', 'title'},
    'div': {'class', 'id', 'style', 'data-*'},
    'span': {'class', 'id', 'style', 'data-*'},
    'table': {'class', 'id'},
    'button': {'onclick', 'disabled'},
    'svg': {'viewBox', 'xmlns', 'width', 'height', 'class'},
}

# Valores de estilo CSS permitidos (whitelist de propriedades)
ALLOWED_CSS_PROPERTIES = {
    'color', 'background-color', 'background', 'font-weight', 'font-size',
    'text-align', 'margin', 'padding', 'border', 'border-radius',
    'display', 'flex', 'align-items', 'justify-content', 'gap',
    'width', 'height', 'max-width', 'min-width', 'max-height', 'min-height',
    'opacity', 'transform', 'transition',
}

# Padrões perigosos em CSS (evitar)
DANGEROUS_CSS_PATTERNS = [
    r'javascript:',
    r'expression\s*\(',
    r'behavior\s*:',
    r'@import',
    r'binding:',
]


class HTMLSanitizer:
    """Sanitizador de HTML robusto."""
    
    def __init__(self, allow_external_urls: bool = False):
        self.allow_external_urls = allow_external_urls
        self.errors: List[str] = []
    
    def sanitize(self, html: str) -> str:
        """
        Sanitiza HTML de entrada.
        
        Args:
            html: HTML a sanitizar
        
        Returns:
            HTML sanitizado
        
        Exemplo:
            sanitizer = HTMLSanitizer()
            clean_html = sanitizer.sanitize(user_html)
        """
        
        self.errors = []
        
        if not html or not isinstance(html, str):
            return ""
        
        # 1. Remover scripts e outros elementos perigosos
        html = self._remove_dangerous_elements(html)
        
        # 2. Usar parser para validar estrutura
        parser = _SafeHTMLParser(
            allowed_tags=ALLOWED_TAGS,
            allowed_attributes=ALLOWED_ATTRIBUTES,
            allow_external_urls=self.allow_external_urls,
        )
        
        try:
            parser.feed(html)
            result = parser.get_safe_html()
        except Exception as e:
            logger.warning(f"Erro ao fazer parse HTML: {e}")
            result = self._escape_html(html)
        
        return result
    
    def sanitize_json_html(self, data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Sanitiza valores HTML em dicionário (recurso frequente em config).
        
        Args:
            data: Dicionário com potencial HTML
        
        Returns:
            Dicionário com HTML sanitizado
        """
        
        sanitized = {}
        
        for key, value in data.items():
            if isinstance(value, str) and self._looks_like_html(value):
                sanitized[key] = self.sanitize(value)
            elif isinstance(value, dict):
                sanitized[key] = self.sanitize_json_html(value)
            elif isinstance(value, list):
                sanitized[key] = [
                    self.sanitize(v) if isinstance(v, str) and self._looks_like_html(v) else v
                    for v in value
                ]
            else:
                sanitized[key] = value
        
        return sanitized
    
    def _remove_dangerous_elements(self, html: str) -> str:
        """Remove elementos claramente perigosos."""
        
        # Remove <script>, <style>, <iframe>, <object>, <embed>, <form>, <input>
        dangerous = ['script', 'style', 'iframe', 'object', 'embed', 'form', 'input']
        
        for tag in dangerous:
            html = re.sub(
                f'<{tag}[^>]*>.*?</{tag}>',
                '',
                html,
                flags=re.IGNORECASE | re.DOTALL
            )
            html = re.sub(
                f'<{tag}[^>]*/?>' ,
                '',
                html,
                flags=re.IGNORECASE
            )
        
        # Remove atributos perigosos
        html = re.sub(r'\s+on\w+\s*=\s*["\'][^"\']*["\']', '', html, flags=re.IGNORECASE)
        html = re.sub(r'\s+on\w+\s*=\s*\{[^}]*\}', '', html, flags=re.IGNORECASE)
        
        return html
    
    def _looks_like_html(self, text: str) -> bool:
        """Detecta se texto parece conter HTML."""
        return bool(re.search(r'<[a-z][^>]*>', text, re.IGNORECASE))
    
    def _escape_html(self, text: str) -> str:
        """Escapa caracteres HTML perigosos."""
        replacements = {
            '&': '&amp;',
            '<': '&lt;',
            '>': '&gt;',
            '"': '&quot;',
            "'": '&#39;',
        }
        
        for char, escaped in replacements.items():
            text = text.replace(char, escaped)
        
        return text


class _SafeHTMLParser(HTMLParser):
    """Parser customizado que constrói HTML seguro."""
    
    def __init__(self, allowed_tags: Set[str], allowed_attributes: Dict[str, Set[str]], 
                 allow_external_urls: bool = False):
        super().__init__()
        self.allowed_tags = allowed_tags
        self.allowed_attributes = allowed_attributes
        self.allow_external_urls = allow_external_urls
        self.output = []
        self.tag_stack = []
    
    def handle_starttag(self, tag: str, attrs):
        """Processa tag de abertura."""
        
        # Tags não permitidas
        if tag.lower() not in self.allowed_tags:
            return
        
        # Filtrar atributos
        safe_attrs = []
        
        for attr_name, attr_value in attrs:
            if not self._is_safe_attribute(tag, attr_name, attr_value):
                continue
            
            safe_attrs.append((attr_name, attr_value))
        
        # Construir tag
        attrs_str = ' '.join(
            f'{k}="{self._escape_html(str(v))}"'
            for k, v in safe_attrs
        ) if safe_attrs else ''
        
        if attrs_str:
            self.output.append(f'<{tag} {attrs_str}>')
        else:
            self.output.append(f'<{tag}>')
        
        self.tag_stack.append(tag.lower())
    
    def handle_endtag(self, tag: str):
        """Processa tag de fechamento."""
        
        if tag.lower() in self.allowed_tags and self.tag_stack and self.tag_stack[-1] == tag.lower():
            self.output.append(f'</{tag}>')
            self.tag_stack.pop()
    
    def handle_data(self, data: str):
        """Processa data/text."""
        # Escapar entidades perigosas
        self.output.append(self._escape_html(data))
    
    def handle_entity(self, name: str):
        """Processa entidade HTML."""
        self.output.append(f'&{name};')
    
    def _is_safe_attribute(self, tag: str, attr_name: str, attr_value: str) -> bool:
        """Valida se atributo é seguro."""
        
        # Sempre rejeitar atributos on*
        if attr_name.lower().startswith('on'):
            return False
        
        # Permitir data-* e class/id para qualquer tag
        if attr_name.lower().startswith('data-') or attr_name.lower() in {'class', 'id', 'style'}:
            if attr_name.lower() == 'style':
                return self._is_safe_css(attr_value)
            return True
        
        # Check allowed attributes por tag
        if tag.lower() in self.allowed_attributes:
            attr_pattern = self.allowed_attributes[tag.lower()]
            
            # Suporta data-* genericamente
            if any(p.endswith('-*') for p in attr_pattern):
                return True
            
            if attr_name.lower() not in attr_pattern:
                return False
        
        # Validar URLs específicos
        if attr_name.lower() in {'href', 'src'}:
            return self._is_safe_url(attr_value)
        
        return True
    
    def _is_safe_url(self, url: str) -> bool:
        """Valida URL é segura."""
        
        url = url.strip().lower()
        
        # Rejeitar javascript:, data:, etc
        dangerous_protocols = {'javascript:', 'data:', 'vbscript:', 'file:'}
        
        for protocol in dangerous_protocols:
            if url.startswith(protocol):
                return False
        
        # Permitir apenas HTTP(S) externo ou caminhos relativos
        if url.startswith('http'):
            return self.allow_external_urls
        
        # Caminhos relativos são OK
        return True
    
    def _is_safe_css(self, css_string: str) -> bool:
        """Valida se CSS é seguro."""
        
        # Verificar padrões perigosos
        for pattern in DANGEROUS_CSS_PATTERNS:
            if re.search(pattern, css_string, re.IGNORECASE):
                return False
        
        # Limite basicamente na sintaxe
        return len(css_string) < 5000
    
    def _escape_html(self, text: str) -> str:
        """Escapa HTML."""
        return (text
                .replace('&', '&amp;')
                .replace('<', '&lt;')
                .replace('>', '&gt;')
                .replace('"', '&quot;')
                .replace("'", '&#39;'))
    
    def get_safe_html(self) -> str:
        """Retorna HTML sanitizado."""
        return ''.join(self.output)


def sanitize_html(html: str) -> str:
    """Interface simples para sanitizar HTML."""
    if not config.get_bool('features.enable_html_sanitization', True):
        return html
    
    sanitizer = HTMLSanitizer()
    return sanitizer.sanitize(html)


def sanitize_dashboard_html(dashboard_dict: Dict[str, Any]) -> Dict[str, Any]:
    """Sanitiza todo HTML em um dicionário de dashboard."""
    sanitizer = HTMLSanitizer()
    return sanitizer.sanitize_json_html(dashboard_dict)


if __name__ == "__main__":
    # Teste
    test_html = """
    <div onclick="alert('XSS')">
        <h1>Seja bem-vindo</h1>
        <img src="x" onerror="alert('XSS')">
        <p style="color: red; background: url(javascript:alert('XSS'))">Texto seguro</p>
        <script>alert('XSS')</script>
        <a href="javascript:void(0)">Link malicioso</a>
        <a href="/relatorio">Link seguro</a>
    </div>
    """
    
    sanitizer = HTMLSanitizer()
    clean = sanitizer.sanitize(test_html)
    print("Original HTML:")
    print(test_html)
    print("\n✅ HTML Sanitizado:")
    print(clean)
