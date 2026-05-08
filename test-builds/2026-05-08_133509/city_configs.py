"""
Configurações centralizadas de empresas/cidades.
Usado por email_dispatch_module.py e email_merge_module.py para evitar duplicação.
"""
from dataclasses import dataclass
from typing import Optional, Dict
import unicodedata


def _normalize_text(text: str) -> str:
    """Remove acentos e converte para maiúsculas para comparação."""
    nfd_form = unicodedata.normalize('NFD', text)
    return ''.join(c for c in nfd_form if unicodedata.category(c) != 'Mn').upper()

@dataclass
class CompanyConfig:
    """Configuração dinâmica para empresas/cidades."""
    name: str
    cnpj: str
    address: str
    phone: str
    whatsapp: str
    website: str
    app: str
    app_link: str
    hours: str
    city: str = ""
    state: str = ""
    logo_url: Optional[str] = None
    primary_color: str = "#0D1F35"
    secondary_color: str = "#F27D16"

    @classmethod
    def from_dict(cls, data: dict) -> "CompanyConfig":
        return cls(
            name=data.get("name", ""),
            cnpj=data.get("cnpj", ""),
            address=data.get("address", ""),
            phone=data.get("phone", ""),
            whatsapp=data.get("whatsapp", ""),
            website=data.get("website", ""),
            app=data.get("app", ""),
            app_link=data.get("app_link", ""),
            hours=data.get("hours", ""),
            city=data.get("city", ""),
            state=data.get("state", ""),
            logo_url=data.get("logo_url"),
            primary_color=data.get("primary_color", "#0D1F35"),
            secondary_color=data.get("secondary_color", "#F27D16"),
        )

    def to_merge_dict(self) -> Dict[str, str]:
        """Converte CompanyConfig para dict de merge (compatível com email_merge_module)."""
        return {
            "EMPRESA_NOME": self.name,
            "EMPRESA_NOME_CURTO": self.name.split()[0],
            "EMPRESA_CNPJ": self.cnpj,
            "EMPRESA_ENDERECO": self.address,
            "EMPRESA_0800": self.phone,
            "EMPRESA_0800_LINK": self.phone.replace(" ", "").replace("-", ""),
            "EMPRESA_SITE": self.website,
            "EMPRESA_WHATSAPP": self.whatsapp,
            "EMPRESA_WHATSAPP_LINK": self.whatsapp.replace("(", "").replace(")", "").replace(" ", "").replace("-", ""),
            "EMPRESA_MUNICIPIO": f"{self.city}/{self.state}",
            "EMPRESA_APP": self.app,
            "EMPRESA_APP_LINK": self.app_link,
            "COR_PRIMARIA": self.primary_color,
            "COR_SECUNDARIA": self.secondary_color,
            "COR_TERCIARIA": "#8FA8B8",
            "EMPRESA_HORARIO": self.hours + ("." if not self.hours.endswith(".") else ""),
        }


# Configurações de todas as cidades/empresas
DEFAULT_COMPANIES: Dict[str, CompanyConfig] = {
    "Itapoá": CompanyConfig(
        name="ITAPOÁ SANEAMENTO LTDA",
        cnpj="CNPJ 16.920.256/0001-57",
        address="RUA 745 LINDÓIA, 328 – ITAPEMA DO NORTE – BALN. PÉROLA – ITAPOÁ – SC",
        phone="0800 643 2750",
        whatsapp="(47) 99278-0310",
        website="www.itapoasaneamento.com.br",
        app="Centro Sul Concessões",
        app_link="https://play.google.com/store/apps/details?id=com.waterfy.agenciavirtual.centrosul",
        hours="de segunda a sexta-feira, das 8h às 16h30",
        city="Itapoá",
        state="SC",
        primary_color="#002B5B",
        secondary_color="#F58220",
        logo_url=None,
    ),
    "Gravatal": CompanyConfig(
        name="GRAVATAL SANEAMENTO",
        cnpj="CNPJ 29.532.193/0001-03",
        address="Rua Engenheiro Annes Gualberto, 85 - Centro<br>Gravatal - Santa Catarina<br>CEP: 88735-000",
        phone="0800 942 9080",
        whatsapp="(48) 99132-6926",
        website="www.gravatalsaneamento.com.br",
        app="Centro Sul Concessões",
        app_link="https://play.google.com/store/apps/details?id=com.waterfy.agenciavirtual.centrosul",
        hours="de segunda a sexta-feira, das 8h às 17h",
        city="Gravatal",
        state="SC",
        primary_color="#0487D9",
        secondary_color="#16B4F2",
        logo_url=None,
    ),
    "Balneário Gaivota": CompanyConfig(
        name="BALNEARIO GAIVOTA LTDA",
        cnpj="CNPJ 30.458.930/0001-54",
        address="Avenida Santa Catarina, 402 - Centro<br>Balneário Gaivota - Santa Catarina<br>CEP: 88955-000",
        phone="0800 942 5572",
        whatsapp="(48) 99173-5908",
        website="www.gaivotasaneamento.com.br",
        app="Centro Sul Concessões",
        app_link="https://play.google.com/store/apps/details?id=com.waterfy.agenciavirtual.centrosul",
        hours="de segunda a sexta-feira, das 8h às 16h30",
        city="Balneário Gaivota",
        state="SC",
        primary_color="#0487D9",
        secondary_color="#16B4F2",
        logo_url=None,
    ),
    "Sombrio": CompanyConfig(
        name="Sombrio Saneamento SPE SA",
        cnpj="CNPJ 39.673.029/0001-70",
        address="Avenida Nereu Ramos, nº 30, sala 02<br>Centro - Sombrio - Santa Catarina<br>CEP: 88960-000",
        phone="0800 048 0820",
        whatsapp="(48) 99183-1580",
        website="https://sombriosaneamento.com.br",
        app="Centro Sul Concessões",
        app_link="https://play.google.com/store/apps/details?id=com.waterfy.agenciavirtual.centrosul",
        hours="de segunda a sexta-feira, das 8h às 16h30",
        city="Sombrio",
        state="SC",
        primary_color="#0487D9",
        secondary_color="#16B4F2",
        logo_url=None,
    ),
    "São Gabriel": CompanyConfig(
        name="São Gabriel Saneamento",
        cnpj="CNPJ 15.186.494/0001-18",
        address="Rua Andrade Neves, 339<br>Praça Fernando Abbott - Centro<br>São Gabriel - RS",
        phone="0800 642 3003",
        whatsapp="(55) 99651-2706",
        website="https://www.sgssa.com.br",
        app="Centro Sul Concessões",
        app_link="https://av-saogabriel.waterfy.net/sign_up",
        hours="de segunda a sexta-feira, das 8h às 16h30",
        city="São Gabriel",
        state="RS",
        primary_color="#7CA865",
        secondary_color="#008DCD",
        logo_url=None,
    ),
    "Nova Cidade": CompanyConfig(
        name="NOME DA EMPRESA",
        cnpj="CNPJ XX.XXX.XXX/XXXX-XX",
        address="Endereço da Empresa",
        phone="0800 XXX XXXX",
        whatsapp="(XX) XXXXX-XXXX",
        website="www.exemplo.com.br",
        app="Nome do App",
        app_link="https://play.google.com/store/apps/details?id=SEU_APP_ID_AQUI",
        hours="de segunda a sexta-feira, das Xh às Xh",
        city="Nova Cidade",
        state="UF",
        primary_color="#000000",
        secondary_color="#FFFFFF",
        logo_url=None,
    ),
}


def get_city_config(city_name: str) -> Optional[CompanyConfig]:
    """
    Busca configuração de cidade.
    Tenta correspondência exata primeiro, depois case-insensitive com remoção de acentos.
    """
    # Tentativa exata
    if city_name in DEFAULT_COMPANIES:
        return DEFAULT_COMPANIES[city_name]
    
    # Tentativa case-insensitive com remoção de acentos
    city_normalized = _normalize_text(city_name)
    for key, config in DEFAULT_COMPANIES.items():
        if (_normalize_text(key) == city_normalized or 
            _normalize_text(config.city) == city_normalized):
            return config
    
    return None
