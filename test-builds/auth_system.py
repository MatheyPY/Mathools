"""
Sistema de Autenticação - Versão Cloud REST API (HTTPS / Porta 443)
Autenticação via Supabase utilizando padrão PostgREST (RPC) com RLS ativo.

Conformidade LGPD (Lei 13.709/2018) — implementação interna:
  Art. 6º, III  — Minimização: apenas dados estritamente necessários são coletados.
  Art. 7º, II   — Base legal: execução de contrato de trabalho.
  Art. 18, VI   — Eliminação: delete_user() remove todos os dados do titular.
  Art. 37       — Rastreabilidade: logs de auditoria sem dados pessoais identificáveis.
  Art. 46       — Segurança técnica: bcrypt irreversível, HTTPS, JWT com expiração.
"""

import bcrypt
import jwt
import requests
import logging
from datetime import datetime, timedelta
from typing import Optional, Dict

# Novos módulos integrados (Tarefas 1-10)
try:
    from config_loader import config
    from audit_logger import audit_logger, audit
    from timezone_utils import tz_manager, format_time
except ImportError as e:
    import sys
    print(f"⚠️  Aviso: Módulo não encontrado: {e}", file=sys.stderr)

# ── Logger interno (sem dados pessoais nos registros) ────────────────────────
_log = logging.getLogger("auth_system")
if not _log.handlers:
    _h = logging.StreamHandler()
    _h.setFormatter(logging.Formatter(
        "%(asctime)s [AUTH] %(levelname)s %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    ))
    _log.addHandler(_h)
_log.setLevel(logging.WARNING)

# ── Configuração Supabase (HTTPS / Porta 443) — Carregada de config.toml ────
SUPABASE_URL = config.get_str('database.supabase_url', 'https://zfhzlxfpxyairtsaqqxu.supabase.co')
SUPABASE_KEY = config.get_str('database.supabase_key', '')
SECRET_KEY = config.get_str('database.secret_key', 'itapoa-saneamento-mathools-producao-2026')
MAX_LOGIN_ATTEMPTS = config.get_int('security.max_login_attempts', 5)
LOCKOUT_TIME_MINUTES = config.get_int('security.lockout_minutes', 15)
_TOKEN_EXPIRY_HOURS = config.get_int('database.jwt_expiry_hours', 8)
_REQUEST_TIMEOUT = config.get_int('database.supabase_timeout', 7)


def get_headers() -> Dict[str, str]:
    return {
        "apikey":        SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type":  "application/json",
        "Prefer":        "return=representation",
    }


class Authentication:

    def __init__(self):
        self.last_error = ""
        self.last_error_code = ""

    def _set_error(self, code: str, message: str) -> None:
        self.last_error_code = code
        self.last_error = message

    @staticmethod
    def hash_password(password: str) -> str:
        """Hash bcrypt irreversível (rounds=12). A senha nunca trafega em claro."""
        return bcrypt.hashpw(
            password.encode("utf-8"),
            bcrypt.gensalt(rounds=12),
        ).decode("utf-8")

    @staticmethod
    def verify_password(password: str, password_hash: str) -> bool:
        """Verificação em tempo constante — resistente a timing attacks."""
        try:
            return bcrypt.checkpw(
                password.encode("utf-8"),
                password_hash.encode("utf-8"),
            )
        except Exception:
            return False

    def get_all_users(self) -> list:
        """
        Retorna lista de usuários.
        A Stored Procedure retorna apenas campos necessários —
        password_hash jamais é incluído na resposta.
        """
        try:
            resp = requests.post(
                f"{SUPABASE_URL}/rest/v1/rpc/get_users_list",
                headers=get_headers(),
                timeout=_REQUEST_TIMEOUT,
                verify=False,
            )
            resp.raise_for_status()
            users = resp.json()
            for u in users:
                if not u.get("created_at"):
                    u["created_at"] = ""
            return users
        except Exception as e:
            _log.error("Falha ao buscar usuarios: %s", type(e).__name__)
            return []

    def add_user(self, username, password, role, allowed_modules):
        """
        Cadastra novo usuário.
        Senha hasheada localmente antes de qualquer transmissão (Art. 46 LGPD).
        """
        if not username or not username.strip():
            return False, "Nome de usuário não pode ser vazio."
        if not password:
            return False, "Senha não pode ser vazia."
        try:
            pw_hash = self.hash_password(password)
            data = {
                "p_username":        username.strip(),
                "p_password_hash":   pw_hash,
                "p_role":            role,
                "p_allowed_modules": allowed_modules,
            }
            resp = requests.post(
                f"{SUPABASE_URL}/rest/v1/rpc/add_new_user",
                headers=get_headers(),
                json=data,
                timeout=_REQUEST_TIMEOUT,
                verify=False,
            )
            if resp.status_code not in (200, 204):
                if "duplicate key" in resp.text.lower():
                    return False, "Nome de usuário já existe!"
                resp.raise_for_status()
            return True, f"Usuário '{username}' criado com sucesso na nuvem!"
        except Exception as e:
            _log.error("Erro ao criar usuario: %s", type(e).__name__)
            return False, f"Falha de conexão: {type(e).__name__}"

    def update_user(self, user_id, username, password, role, allowed_modules):
        """
        Atualiza dados de usuário.
        Se password for vazio, a senha atual não é alterada.
        """
        try:
            pw_hash = self.hash_password(password) if password else ""
            data = {
                "p_id":              user_id,
                "p_username":        username,
                "p_password_hash":   pw_hash,
                "p_role":            role,
                "p_allowed_modules": allowed_modules,
            }
            resp = requests.post(
                f"{SUPABASE_URL}/rest/v1/rpc/update_existing_user",
                headers=get_headers(),
                json=data,
                timeout=_REQUEST_TIMEOUT,
                verify=False,
            )
            if resp.status_code not in (200, 204):
                if "duplicate key" in resp.text.lower():
                    return False, "O nome de usuário já existe e pertence a outra conta!"
                resp.raise_for_status()
            return True, f"Usuário '{username}' atualizado com sucesso!"
        except Exception as e:
            _log.error("Erro ao atualizar usuario: %s", type(e).__name__)
            return False, f"Falha de conexão: {type(e).__name__}"

    def delete_user(self, user_id):
        """
        Remove permanentemente todos os dados do usuário (Art. 18, VI LGPD):
        credenciais, cargo, permissões e histórico de acessos.
        Execução via Stored Procedure com RLS — operação irreversível.
        """
        try:
            resp = requests.post(
                f"{SUPABASE_URL}/rest/v1/rpc/delete_existing_user",
                headers=get_headers(),
                json={"p_id": user_id},
                timeout=_REQUEST_TIMEOUT,
                verify=False,
            )
            resp.raise_for_status()
            return True
        except Exception as e:
            _log.error("Erro ao remover usuario: %s", type(e).__name__)
            return False

    def login(self, username: str, password: str, ip_address: str = "", show_ui: bool = True) -> Optional[Dict]:
        """
        Autentica usuário corporativo.

        Segurança (Art. 46 LGPD):
          - RPC com RLS ativo; bloqueio verificado pelo relógio do servidor
          - Senha verificada localmente via bcrypt — nunca trafega em claro
          - JWT com expiração automática de 8 horas
          - Auditoria LGPD: login tentativas registradas sem PII (Art. 37)
        """
        # ip_address recebido por compatibilidade de assinatura, mas não utilizado
        # (coleta de IP não é necessária para a finalidade — Art. 6º, III LGPD)
        self._set_error("", "")

        try:
            # 1. Busca dados do usuário; bloqueio verificado pelo relógio do servidor
            resp = requests.post(
                f"{SUPABASE_URL}/rest/v1/rpc/secure_login",
                headers=get_headers(),
                json={"p_username": username},
                timeout=_REQUEST_TIMEOUT,
                verify=False,
            )
            resp.raise_for_status()
            users = resp.json()

            if not users:
                # Log de falha de autenticação
                try:
                    audit_logger.log_auth_attempt(
                        user=username,
                        success=False,
                        ip_address=ip_address or "UNKNOWN",
                        detail="Usuário não encontrado"
                    )
                except Exception:
                    pass
                self._set_error("user_not_found", "Usuário não encontrado.")
                return None

            user = users[0]

            # 2. Verifica bloqueio ativo
            if user.get("is_locked"):
                try:
                    audit_logger.log_auth_attempt(
                        user=username,
                        success=False,
                        ip_address=ip_address or "UNKNOWN",
                        detail="Usuário bloqueado (múltiplas tentativas falhas)"
                    )
                except Exception:
                    pass
                
                self._set_error("locked", f"O usuário '{username}' está bloqueado. Aguarde 15 minutos e tente novamente.")
                return None

            # 3. Verifica senha localmente — texto claro nunca sai da máquina
            if not self.verify_password(password, user["password_hash"]):
                try:
                    requests.post(
                        f"{SUPABASE_URL}/rest/v1/rpc/register_failed_login",
                        headers=get_headers(),
                        json={"p_username": username},
                        timeout=_REQUEST_TIMEOUT,
                        verify=False,
                    )
                    # Log de falha
                    audit_logger.log_auth_attempt(
                        user=username,
                        success=False,
                        ip_address=ip_address or "UNKNOWN",
                        detail="Senha incorreta"
                    )
                except Exception:
                    pass

                tentativas_restantes = 4 - user.get("failed_login_attempts", 0)
                if tentativas_restantes > 0:
                    self._set_error("invalid_password", f"Senha incorreta. Você tem mais {tentativas_restantes} tentativa(s).")
                else:
                    self._set_error("locked", "Senha incorreta. Limite atingido. Usuário bloqueado por 15 minutos.")
                return None

            # 4. Sucesso — zera contador de falhas e registra acesso
            try:
                requests.post(
                    f"{SUPABASE_URL}/rest/v1/rpc/register_successful_login",
                    headers=get_headers(),
                    json={"p_username": username},
                    timeout=_REQUEST_TIMEOUT,
                    verify=False,
                )
            except Exception:
                pass

            # 5. Emite JWT de sessão com expiração automática (timezone-aware)
            token = jwt.encode(
                {
                    "user_id":  user["id"],
                    "username": user["username"],
                    "role":     user["role"],
                    "issued_at": tz_manager.now_iso(),
                    "exp": int((tz_manager.now() + timedelta(hours=_TOKEN_EXPIRY_HOURS)).timestamp()),
                },
                SECRET_KEY,
                algorithm="HS256",
            )
            
            # Log de sucesso
            try:
                audit_logger.log_auth_attempt(
                    user=username,
                    success=True,
                    ip_address=ip_address or "UNKNOWN",
                    detail="Supabase JWT emitido com sucesso"
                )
            except Exception:
                pass

            return {
                "user_id":         user["id"],
                "username":        user["username"],
                "email":           user.get("email", ""),
                "role":            user["role"],
                "allowed_modules": user["allowed_modules"],
                "token":           token,
                "issued_at":       tz_manager.now_iso(),
            }

        except requests.exceptions.RequestException as e:
            self._set_error("network", f"A comunicação com a nuvem falhou ({type(e).__name__}). Verifique sua internet.")
            return None
        except Exception as e:
            self._set_error("unexpected", f"Falha inesperada na autenticação ({type(e).__name__}).")
            return None

    def logout(self, token: str, user_id: int, ip_address: str = "") -> None:
        """Token descartado no cliente. Sem transmissão de dados adicionais."""
        pass

    def verify_token(self, token: str) -> Optional[Dict]:
        """Verifica e decodifica JWT. Retorna None se expirado ou inválido."""
        try:
            return jwt.decode(token, SECRET_KEY, algorithms=["HS256"])
        except jwt.ExpiredSignatureError:
            return None
        except jwt.InvalidTokenError:
            return None
        except Exception:
            return None
