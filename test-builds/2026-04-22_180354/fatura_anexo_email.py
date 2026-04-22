"""
Resolve o caminho do PDF de fatura para anexar em e-mails (mesma regra do disparo WhatsApp).
Prioridade: coluna de arquivo customizado → download Waterfy (MATRICULA + ID_FATURA) → concatenação.
"""
from __future__ import annotations

import os
import re
import shutil
import concurrent.futures
from typing import Any, Dict, Optional, Tuple

COLUNAS_ARQUIVO_CUSTOM = [
    "ARQUIVO_PDF",
    "ARQUIVO",
    "PDF",
    "CAMINHO_PDF",
    "DOCUMENTO",
    "ARQUIVO_DOC",
    "FILE",
    "PATH_PDF",
    "ANEXO",
    "ARQUIVO_ENVIO",
]


def _row_upper_keys(row: Dict[str, Any]) -> Dict[str, str]:
    out: Dict[str, str] = {}
    for k, v in row.items():
        key = str(k).strip().upper()
        if v is None or (isinstance(v, float) and str(v) == "nan"):
            out[key] = ""
        else:
            out[key] = str(v).strip()
    return out


def _get_cell(row_up: Dict[str, str], *aliases: str) -> str:
    for a in aliases:
        key = a.strip().upper()
        if key in row_up and row_up[key]:
            return row_up[key].strip()
    return ""


def resolve_fatura_anexo_para_email(
    row: Dict[str, Any],
    *,
    linha: int,
    baixar_faturas: bool,
    usuario_waterfy: str,
    senha_waterfy: str,
    proteger_pdf: bool,
    destino_dir: str,
) -> Tuple[Optional[str], str]:
    """
    Retorna (caminho_absoluto_do_pdf_ou_None, mensagem_aviso).
    mensagem_aviso é não vazia apenas em falhas ou avisos úteis.
    """
    row_up = _row_upper_keys(row)

    for col in COLUNAS_ARQUIVO_CUSTOM:
        if col in row_up:
            p = row_up[col].strip()
            if p and os.path.isfile(p):
                return os.path.abspath(p), ""

    if not baixar_faturas:
        return None, ""

    try:
        import mathtools_1_0 as mt
    except Exception as ex:
        return None, f"Não foi possível carregar mathtools: {ex}"

    if not getattr(mt, "HAS_FATURA_DOWNLOADER", False):
        return None, "Download de faturas (Waterfy) não disponível neste ambiente."

    matricula = _get_cell(row_up, "MATRICULA", "MATRÍCULA")
    raw_id = _get_cell(row_up, "ID_FATURA")
    id_faturas = mt._split_id_faturas(raw_id)

    cpf = re.sub(
        r"\D",
        "",
        _get_cell(row_up, "CPF CONSUMIDOR", "CPF_CONSUMIDOR", "CPF"),
    )
    cnpj = re.sub(
        r"\D",
        "",
        _get_cell(row_up, "CNPJ CONSUMIDOR", "CNPJ_CONSUMIDOR", "CNPJ"),
    )
    doc_senha = cpf or cnpj

    os.makedirs(destino_dir, exist_ok=True)
    linha_dir = os.path.join(destino_dir, f"linha_{linha}")
    os.makedirs(linha_dir, exist_ok=True)

    if not matricula or not id_faturas:
        return None, "Matrícula ou ID_FATURA ausente — anexo não gerado."

    if len(id_faturas) == 1:
        id_fatura = id_faturas[0]
        usar_encrypt = bool(proteger_pdf and doc_senha)
        fatura_path = mt.baixar_fatura_cliente(
            matricula,
            id_fatura,
            usuario_waterfy,
            senha_waterfy,
            linha_dir,
            cpf_consumidor=doc_senha or None,
            encrypt=usar_encrypt,
        )
        if fatura_path and os.path.isfile(fatura_path):
            final_path = os.path.join(
                linha_dir,
                f"Segunda_Via_Faturas_Matricula_{matricula}_L{linha}.pdf",
            )
            shutil.copy2(fatura_path, final_path)
            return os.path.abspath(final_path), (
                ""
                if usar_encrypt or not proteger_pdf
                else "Senha no PDF solicitada, mas CPF/CNPJ do consumidor não encontrado na linha."
            )
        return None, "Falha ao baixar a fatura (Waterfy)."

    downloaded_paths: list[str] = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=min(5, len(id_faturas))) as executor:
        futures_dict = {
            executor.submit(
                mt.baixar_fatura_cliente,
                matricula,
                id_fat,
                usuario_waterfy,
                senha_waterfy,
                linha_dir,
                cpf_consumidor=None,
                encrypt=False,
            ): id_fat for id_fat in id_faturas
        }
        results_dict = {}
        for future in concurrent.futures.as_completed(futures_dict):
            try:
                path = future.result()
                if path and os.path.isfile(path):
                    results_dict[futures_dict[future]] = path
            except Exception:
                pass

    for id_fat in id_faturas:
        if id_fat in results_dict:
            downloaded_paths.append(results_dict[id_fat])

    if not downloaded_paths:
        return None, "Nenhuma fatura foi baixada para concatenação."

    merged_path = os.path.join(
        linha_dir,
        f"Segunda_Via_Faturas_Matricula_{matricula}_L{linha}.pdf",
    )
    merge_pwd = (doc_senha if proteger_pdf else None) or None
    aviso = ""
    if proteger_pdf and not doc_senha:
        aviso = "PDF concatenado sem senha: CPF/CNPJ do consumidor não encontrado na linha."

    if mt._merge_pdfs(downloaded_paths, merged_path, password=merge_pwd):
        return os.path.abspath(merged_path), aviso

    try:
        shutil.copy2(downloaded_paths[0], merged_path)
        return os.path.abspath(merged_path), (
            aviso + " " if aviso else ""
        ) + "Não foi possível concatenar todos os PDFs. Verifique se a biblioteca pypdf está instalada; anexando apenas o primeiro."
    except Exception:
        return os.path.abspath(downloaded_paths[0]), (
            aviso + " " if aviso else ""
        ) + "Concatenação falhou; anexando o primeiro PDF baixado."
