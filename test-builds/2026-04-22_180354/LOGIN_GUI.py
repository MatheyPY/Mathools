#!/usr/bin/env python3
import asyncio
import ctypes
from datetime import datetime
import os
from pathlib import Path
import struct
import sys

import flet as ft

try:
    import launcher_gui as _launcher_gui
except Exception:
    _launcher_gui = None


def resource_path(relative_path: str) -> str:
    if getattr(sys, "frozen", False):
        base = sys._MEIPASS
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, relative_path)


try:
    from updater import check_update_async, apply_update
    _HAS_UPDATER = True
except ImportError:
    _HAS_UPDATER = False

try:
    from auth_system import Authentication
except ImportError as _e:
    print(f"ERRO CRITICO: modulo 'auth_system' nao encontrado. Detalhe: {_e}", file=sys.stderr)
    sys.exit(1)


APP_FONT = "Helvetica"
BG_NAVY = "#012340"
BG_NAVY_LIGHT = "#154F86"
BG_NAVY_FOCUS = "#1C6BB5"
PANEL_SURFACE = "#0A2E52"
PANEL_SURFACE_SOFT = "#113B67"
TXT_WHITE = "#FFFFFF"
TXT_GRAY = "#8A9AB8"
TXT_LABEL = "#A3B4CC"
BRAND_ORANGE = "#F27D16"
BRAND_ORANGE_HOVER = "#E06B00"
ERROR_COLOR = "#FF6B6B"
SUCCESS_COLOR = "#2E7D32"


def _configure_windows_dpi_awareness():
    if sys.platform != "win32":
        return
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        return
    except Exception:
        pass
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


def _greeting():
    hora_atual = datetime.now().hour
    if 5 <= hora_atual < 12:
        return "Bom dia!", "Que o seu dia seja excelente e positivo."
    if 12 <= hora_atual < 18:
        return "Boa tarde!", "Esperamos que a sua tarde esteja rendendo muito."
    return "Boa noite!", "Otimo final de expediente ou um bom turno!"


class LoginWindow:
    def __init__(self):
        self.auth = Authentication()
        self.login_result = None
        self.attempt_count = 0
        self.left_panel_width = 440
        self.image_width = 640
        self.image_height = 638
        self.content_width = 1080
        self.content_height = 638
        self.win_width = 1096
        self.win_height = 662
        self._update_info = None
        self._update_banner_shown = False
        self._page = None
        self._user_field = None
        self._password_field = None
        self._login_button = None
        self._login_button_label = None
        self._error_text = None
        self._update_banner = None
        self._banner_host = None
        self._left_panel = None
        self._right_panel = None
        self._right_stack = None
        self._hero_image = None
        self._hero_overlay = None
        self._hero_edge = None
        self._root_row = None
        self._root_container = None
        self._logo_image = None
        self._logo_block = None
        self._eyebrow_text = None
        self._title_text = None
        self._message_text = None
        self._form_card = None
        self._form_header = None
        self._content_column = None
        self._main_block = None
        self._spacer_before_main = None
        self._spacer_after_banner = None
        self._spacer_after_logo = None
        self._spacer_after_eyebrow = None
        self._spacer_before_form = None
        self._footer_container = None
        self._footer_text = "© 2026 Itapoá Saneamento. Todos os direitos reservados."

        _img_candidates = [
            resource_path("imagem_a_direita_do_loguin.jpg"),
            resource_path("imagem a direita do loguin.jpg"),
            resource_path("imagem_a_direita_do_loguin.png"),
            resource_path("imagem a direita do loguin.png"),
        ]
        self.image_path = next((p for p in _img_candidates if os.path.exists(p)), _img_candidates[0])
        self.image_width, self.image_height = self._read_image_size(self.image_path)
        self.content_width = self.left_panel_width + self.image_width
        self.content_height = max(600, self.image_height)
        self.win_width = self.content_width + 16
        self.win_height = self.content_height + 24

        self.logo_path = resource_path("Itapoa_branca.png")
        self.icon_path = resource_path("Logo-mathey-tk-3.ico")

    @staticmethod
    def _clamp(value, minimum, maximum):
        return max(minimum, min(value, maximum))

    @staticmethod
    def _read_image_size(image_path):
        try:
            with open(image_path, "rb") as f:
                head = f.read(32)
            if head[:8] == b"\x89PNG\r\n\x1a\n":
                return struct.unpack(">II", head[16:24])
        except Exception:
            pass
        return 702, 638

    def _publish_update_info(self, info):
        self._update_info = info if (info and getattr(info, "has_update", False)) else None
        try:
            if _launcher_gui is not None:
                _launcher_gui._LAST_UPDATE_INFO = self._update_info
        except Exception:
            pass

    def _start_update_check(self):
        if not _HAS_UPDATER:
            return
        try:
            check_update_async(self._publish_update_info)
        except Exception:
            pass

    def _set_error(self, message=""):
        if self._error_text is None:
            return
        self._error_text.value = message or ""
        self._error_text.visible = bool(message)
        try:
            if self._page:
                self._page.update()
            else:
                self._error_text.update()
        except Exception:
            pass

    def _set_busy(self, busy: bool):
        if self._login_button is None:
            return
        self._login_button.disabled = busy
        if self._login_button_label is not None:
            self._login_button_label.value = "AUTENTICANDO..." if busy else "ENTRAR"
            self._login_button_label.color = "#E6EEF8" if busy else TXT_WHITE
        try:
            if self._page:
                self._page.update()
            else:
                self._login_button.update()
        except Exception:
            pass

    async def _close_page(self):
        try:
            if self._page and self._page.window:
                await self._page.window.close()
        except Exception:
            pass

    async def _do_login_flow(self, username: str, password: str):
        try:
            result = await asyncio.to_thread(self.auth.login, username, password, "", False)
            if result:
                self.login_result = result
                await self._finish_login_ok()
                return
            auth_msg = getattr(self.auth, "last_error", "") or "Falha na autenticacao."
            auth_code = getattr(self.auth, "last_error_code", "")
            if auth_code == "invalid_password":
                self.attempt_count += 1
            elif auth_code == "locked":
                self.attempt_count = 5
            await self._finish_login_fail(auth_msg)
        except Exception as e:
            await self._finish_login_fail(f"Falha no sistema: {e}")

    async def _finish_login_ok(self):
        self._set_busy(False)
        if self._password_field:
            self._password_field.value = ""
            self._password_field.password = True
        self._set_error("")
        try:
            self._page.update()
        except Exception:
            pass
        await asyncio.sleep(0.1)
        await self._close_page()

    async def _finish_login_fail(self, message: str):
        self._set_busy(False)
        self._set_error(message)
        if self._password_field:
            self._password_field.value = ""
        try:
            self._page.update()
        except Exception:
            pass

    def _handle_login(self, e=None):
        username = (self._user_field.value or "").strip()
        password = self._password_field.value or ""

        if not username or not password:
            self._set_error("Por favor, preencha usuario e senha.")
            try:
                self._page.snack_bar = ft.SnackBar(
                    ft.Text("Preencha usuario e senha.", color=TXT_WHITE),
                    bgcolor=ERROR_COLOR,
                    open=True,
                )
                self._page.update()
            except Exception:
                pass
            return

        self._set_error("")
        self._set_busy(True)
        try:
            self._page.run_task(self._do_login_flow, username, password)
        except Exception as e:
            self._set_busy(False)
            self._set_error(f"Falha ao iniciar login: {e}")

    async def _download_update(self, dialog, status_text, progress_bar, percent_text):
        def _progress(downloaded, total):
            pct = 0
            if total:
                pct = max(0, min(100, int(downloaded / total * 100)))
            progress_bar.value = pct / 100.0
            percent_text.value = f"{pct}%"
            try:
                dialog.update()
            except Exception:
                pass

        ok = False
        try:
            ok = await asyncio.to_thread(apply_update, self._update_info, _progress)
        except Exception:
            ok = False

        if ok:
            status_text.value = "Download concluido. O Mathools sera fechado para concluir a instalacao."
            try:
                dialog.update()
            except Exception:
                pass
            await asyncio.sleep(0.8)
            await self._close_page()
        else:
            status_text.value = "Falha ao baixar a atualizacao. Tente novamente mais tarde."
            try:
                dialog.update()
            except Exception:
                pass

    def _open_update_dialog(self, e=None):
        if not self._update_info or self._page is None:
            return

        progress_bar = ft.ProgressBar(value=0, color="#7CFC00", bgcolor=ft.Colors.with_opacity(0.18, ft.Colors.WHITE))
        percent_text = ft.Text("0%", color=TXT_GRAY, size=12)
        status_text = ft.Text(
            f"Baixando Mathools {self._update_info.latest_version}...",
            color=TXT_WHITE,
            size=13,
        )

        dialog = ft.AlertDialog(
            modal=True,
            bgcolor=BG_NAVY,
            title=ft.Text("Atualizacao disponivel", color=TXT_WHITE, weight=ft.FontWeight.BOLD),
            content=ft.Container(
                width=420,
                content=ft.Column(
                    [
                        ft.Text(
                            f"Versao {self._update_info.latest_version}",
                            color="#7CFC00",
                            size=15,
                            weight=ft.FontWeight.BOLD,
                        ),
                        ft.Text(self._update_info.release_notes or "", color=TXT_GRAY, size=12),
                        ft.Container(height=10),
                        status_text,
                        ft.Container(height=8),
                        progress_bar,
                        ft.Container(height=4),
                        percent_text,
                    ],
                    tight=True,
                ),
            ),
            actions=[],
        )
        self._page.show_dialog(dialog)
        self._page.run_task(self._download_update, dialog, status_text, progress_bar, percent_text)

    def _refresh_update_banner_layout(self, left_width):
        if self._banner_host is None:
            return
        if self._update_banner is not None:
            self._banner_host.controls = [ft.Container(width=max(270, left_width - 88), content=self._update_banner)]
        try:
            self._banner_host.update()
        except Exception:
            pass

    def _apply_responsive_layout(self):
        self._fit_to_client_area()

    async def _center_window(self):
        try:
            if self._page and self._page.window:
                await self._page.window.center()
        except Exception:
            pass

    async def _focus_control(self, control):
        try:
            if control is not None:
                await control.focus()
        except Exception:
            pass

    def _queue_focus(self, control):
        if self._page is None or control is None:
            return
        try:
            self._page.run_task(self._focus_control, control)
        except Exception:
            pass

    def _handle_username_submit(self, e=None):
        self._queue_focus(self._password_field)

    def _fit_to_client_area(self, e=None):
        if (
            self._page is None
            or self._left_panel is None
            or self._right_panel is None
            or self._right_stack is None
            or self._hero_overlay is None
            or self._root_row is None
        ):
            return

        client_w = int(self._page.width or self.content_width)
        client_h = int(self._page.height or self.content_height)
        narrow_w = client_w < 560
        compact_h = client_h < 690
        ultra_compact = client_h < 610
        if narrow_w:
            left_width = client_w
        else:
            left_width = self._clamp(int(client_w * 0.44), 400, self.left_panel_width)
        right_width = max(0, client_w - left_width)

        self._root_row.width = client_w
        self._root_row.height = client_h
        if self._root_container is not None:
            self._root_container.width = client_w
            self._root_container.height = client_h
        self._left_panel.width = left_width
        self._left_panel.height = client_h
        self._right_panel.width = right_width
        self._right_panel.height = client_h
        self._right_panel.visible = right_width > 0
        self._right_stack.width = right_width
        self._right_stack.height = client_h
        self._hero_overlay.width = right_width
        self._hero_overlay.height = client_h
        if self._hero_edge is not None:
            self._hero_edge.height = client_h
        if self._hero_image is not None:
            self._hero_image.width = right_width
            self._hero_image.height = client_h
        if self._logo_block is not None:
            self._logo_block.width = max(220, left_width - (56 if narrow_w else 88))
            self._logo_block.visible = True
        if self._logo_image is not None:
            self._logo_image.height = 60 if ultra_compact else 70 if compact_h else 88
        if self._eyebrow_text is not None:
            self._eyebrow_text.visible = True
            self._eyebrow_text.size = 10 if ultra_compact else 11 if compact_h else 12
        if self._title_text is not None:
            self._title_text.size = 22 if ultra_compact else 25 if compact_h else 30
        if self._message_text is not None:
            self._message_text.size = 10 if ultra_compact else 11 if compact_h else 12
        if self._user_field is not None:
            field_width = max(252, min(316, left_width - (60 if narrow_w else 80)))
            self._user_field.width = field_width
            self._user_field.height = 48 if ultra_compact else 52 if compact_h else 58
            self._user_field.text_size = 13 if ultra_compact else 14 if compact_h else 15
        if self._password_field is not None:
            self._password_field.width = self._user_field.width if self._user_field is not None else max(240, min(300, left_width - 88))
            self._password_field.height = 48 if ultra_compact else 52 if compact_h else 58
            self._password_field.text_size = 13 if ultra_compact else 14 if compact_h else 15
        if getattr(self, "_form_header", None) is not None and self._user_field is not None:
            self._form_header.width = self._user_field.width
        if self._login_button is not None and self._user_field is not None:
            self._login_button.width = self._user_field.width
        if self._form_card is not None:
            self._form_card.border_radius = 22 if ultra_compact else 24 if compact_h else 28
            self._form_card.padding = ft.padding.only(
                left=16 if ultra_compact else 18 if compact_h else 22,
                right=16 if ultra_compact else 18 if compact_h else 22,
                top=14 if ultra_compact else 16 if compact_h else 18,
                bottom=14 if ultra_compact else 16 if compact_h else 18,
            )
        if self._spacer_after_banner is not None:
            self._spacer_after_banner.height = 2 if ultra_compact else 4 if compact_h else 6
        if self._spacer_before_main is not None:
            self._spacer_before_main.height = 10 if ultra_compact else 14 if compact_h else 22
        if self._spacer_after_logo is not None:
            self._spacer_after_logo.height = 10 if ultra_compact else 14 if compact_h else 18
        if self._spacer_after_eyebrow is not None:
            self._spacer_after_eyebrow.height = 4 if ultra_compact else 6 if compact_h else 8
        if self._spacer_before_form is not None:
            self._spacer_before_form.height = 10 if ultra_compact else 12 if compact_h else 14
        if self._footer_container is not None:
            self._footer_container.visible = not narrow_w
        if self._content_column is not None:
            self._content_column.scroll = ft.ScrollMode.AUTO if ultra_compact or narrow_w else ft.ScrollMode.HIDDEN
        if self._login_button is not None:
            self._login_button.style = ft.ButtonStyle(
                bgcolor={"": BRAND_ORANGE, "hovered": BRAND_ORANGE_HOVER, "disabled": "#4A5D78"},
                color=TXT_WHITE,
                shape=ft.RoundedRectangleBorder(radius=24 if ultra_compact else 26 if compact_h else 28),
                text_style=ft.TextStyle(size=14 if ultra_compact else 15 if compact_h else 16, weight=ft.FontWeight.BOLD, font_family=APP_FONT),
                padding=ft.padding.symmetric(vertical=12 if ultra_compact else 14 if compact_h else 18),
            )
        self._left_panel.padding = ft.padding.only(
            left=18 if narrow_w else 28 if ultra_compact else 36 if compact_h else 44,
            right=18 if narrow_w else 28 if ultra_compact else 36 if compact_h else 44,
            top=24 if ultra_compact else 28 if compact_h else 36,
            bottom=10 if ultra_compact else 12 if compact_h else 16,
        )
        self._refresh_update_banner_layout(left_width)
        try:
            self._root_row.update()
        except Exception:
            pass

    async def _poll_update_banner(self):
        while self._page is not None and self.login_result is None:
            if self._update_info and not self._update_banner_shown and self._banner_host is not None:
                self._update_banner_shown = True
                self._update_banner = ft.Container(
                    bgcolor="#163E1E",
                    border_radius=16,
                    padding=ft.padding.symmetric(horizontal=18, vertical=12),
                    content=ft.Row(
                        [
                            ft.Icon(ft.Icons.SYSTEM_UPDATE_ALT, color="#7CFC00", size=18),
                            ft.Text(
                                f"Nova versao {self._update_info.latest_version} disponivel.",
                                color="#E8FFE8",
                                size=12,
                                weight=ft.FontWeight.W_600,
                            ),
                            ft.Container(expand=True),
                            ft.TextButton(
                                "ATUALIZAR",
                                on_click=self._open_update_dialog,
                                style=ft.ButtonStyle(color="#7CFC00"),
                            ),
                        ],
                        alignment=ft.MainAxisAlignment.START,
                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                    ),
                )
                self._banner_host.controls = [ft.Container(width=360, content=self._update_banner)]
                self._banner_host.visible = True
                self._apply_responsive_layout()
            await asyncio.sleep(0.5)

    def _main(self, page: ft.Page):
        self._page = page
        _configure_windows_dpi_awareness()
        page.title = "Mathools 1.0 - Login"
        page.padding = 0
        page.spacing = 0
        page.bgcolor = BG_NAVY
        page.horizontal_alignment = ft.CrossAxisAlignment.STRETCH
        page.scroll = ft.ScrollMode.HIDDEN
        page.theme_mode = ft.ThemeMode.LIGHT
        page.window.width = self.win_width
        page.window.height = self.win_height
        page.window.min_width = self.win_width
        page.window.min_height = self.win_height
        page.window.resizable = False
        page.window.maximizable = False
        page.run_task(self._center_window)
        page.on_resized = self._fit_to_client_area
        try:
            if os.path.exists(self.icon_path):
                page.window.icon = self.icon_path
        except Exception:
            pass
        try:
            if sys.platform == "win32":
                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(u"Mathools.Login")
        except Exception:
            pass

        saudacao, mensagem = _greeting()
        has_side_image = os.path.exists(self.image_path)
        has_logo = os.path.exists(self.logo_path)

        self._banner_host = ft.Column(visible=False, spacing=0)
        self._error_text = ft.Text("", color=ERROR_COLOR, size=12, visible=False)

        self._user_field = ft.TextField(
            label="Usuario",
            width=316,
            border_radius=28,
            filled=True,
            bgcolor=BG_NAVY_LIGHT,
            border_color="#1D5C95",
            focused_border_color=BG_NAVY_FOCUS,
            color=TXT_WHITE,
            label_style=ft.TextStyle(color=TXT_LABEL, size=13, weight=ft.FontWeight.W_600),
            cursor_color=BRAND_ORANGE,
            text_size=15,
            content_padding=ft.padding.symmetric(horizontal=18, vertical=16),
            height=58,
            on_submit=self._handle_username_submit,
        )

        self._password_field = ft.TextField(
            label="Senha",
            width=316,
            border_radius=28,
            filled=True,
            bgcolor=BG_NAVY_LIGHT,
            border_color="#1D5C95",
            focused_border_color=BG_NAVY_FOCUS,
            color=TXT_WHITE,
            label_style=ft.TextStyle(color=TXT_LABEL, size=13, weight=ft.FontWeight.W_600),
            cursor_color=BRAND_ORANGE,
            text_size=15,
            content_padding=ft.padding.symmetric(horizontal=18, vertical=16),
            password=True,
            can_reveal_password=True,
            height=58,
            on_submit=self._handle_login,
        )

        self._login_button_label = ft.Text(
            "ENTRAR",
            color=TXT_WHITE,
            size=16,
            weight=ft.FontWeight.BOLD,
            font_family=APP_FONT,
            text_align=ft.TextAlign.CENTER,
        )
        self._login_button = ft.TextButton(
            content=ft.Row([self._login_button_label], alignment=ft.MainAxisAlignment.CENTER),
            on_click=self._handle_login,
            style=ft.ButtonStyle(
                bgcolor={"": BRAND_ORANGE, "hovered": BRAND_ORANGE_HOVER, "disabled": "#4A5D78"},
                color=TXT_WHITE,
                shape=ft.RoundedRectangleBorder(radius=28),
                padding=ft.padding.symmetric(vertical=18),
            ),
        )
        self._logo_image = ft.Image(src=self.logo_path, fit="contain", height=106) if has_logo else ft.Container(height=106)
        self._logo_block = ft.Container(
            width=self.left_panel_width - 88,
            visible=True,
            content=ft.Row(
                [self._logo_image],
                alignment=ft.MainAxisAlignment.CENTER,
            ),
        )
        self._eyebrow_text = ft.Text(
            "ACESSO AO PAINEL OPERACIONAL",
            color="#8DB8E8",
            size=12,
            weight=ft.FontWeight.W_600,
            font_family=APP_FONT,
            visible=True,
        )
        self._title_text = ft.Text(
            saudacao,
            color=TXT_WHITE,
            size=32,
            weight=ft.FontWeight.BOLD,
            font_family=APP_FONT,
        )
        self._message_text = ft.Text(
            mensagem,
            color=TXT_GRAY,
            size=13,
            font_family=APP_FONT,
        )
        self._form_card = ft.Container(
            bgcolor=PANEL_SURFACE,
            border_radius=28,
            border=ft.border.all(1, ft.Colors.with_opacity(0.08, ft.Colors.WHITE)),
            padding=ft.padding.only(left=22, right=22, top=18, bottom=18),
            content=ft.Column(
                [
                    ft.Row(
                        [
                            ft.Container(
                                width=316,
                                content=ft.Column(
                                    [
                                        ft.Text(
                                            "Identificacao",
                                            color=TXT_WHITE,
                                            size=14,
                                            weight=ft.FontWeight.BOLD,
                                            font_family=APP_FONT,
                                        ),
                                        ft.Text(
                                            "Use seu usuario corporativo e sua senha para entrar no painel.",
                                            color=TXT_GRAY,
                                            size=11,
                                            font_family=APP_FONT,
                                        ),
                                    ],
                                    spacing=2,
                                    tight=True,
                                ),
                            )
                        ],
                        alignment=ft.MainAxisAlignment.CENTER,
                    ),
                    ft.Container(height=10),
                    ft.Row([self._user_field], alignment=ft.MainAxisAlignment.CENTER),
                    ft.Container(height=10),
                    ft.Row([self._password_field], alignment=ft.MainAxisAlignment.CENTER),
                    ft.Container(height=6),
                    self._error_text,
                    ft.Container(height=10),
                    ft.Row(
                        [
                            ft.Container(
                                width=316,
                                content=self._login_button,
                            )
                        ],
                        alignment=ft.MainAxisAlignment.CENTER,
                    ),
                ],
                spacing=0,
                tight=True,
            ),
        )
        self._form_header = self._form_card.content.controls[0].controls[0]
        self._spacer_after_banner = ft.Container(height=8)
        self._spacer_before_main = ft.Container(height=22)
        self._spacer_after_logo = ft.Container(height=18)
        self._spacer_after_eyebrow = ft.Container(height=8)
        self._spacer_before_form = ft.Container(height=16)
        self._footer_container = ft.Container(
            padding=ft.padding.only(top=6),
            visible=True,
            content=ft.Text(
                self._footer_text,
                color=TXT_GRAY,
                size=11,
                text_align=ft.TextAlign.CENTER,
            ),
        )
        self._main_block = ft.Column(
            [
                self._logo_block,
                self._spacer_after_logo,
                self._eyebrow_text,
                self._spacer_after_eyebrow,
                self._title_text,
                self._message_text,
                self._spacer_before_form,
                self._form_card,
                ft.Container(height=18),
                self._footer_container,
            ],
            tight=True,
            alignment=ft.MainAxisAlignment.START,
            horizontal_alignment=ft.CrossAxisAlignment.STRETCH,
            spacing=0,
        )
        self._content_column = ft.Column(
            [
                self._banner_host,
                self._spacer_after_banner,
                self._spacer_before_main,
                self._main_block,
            ],
            expand=True,
            alignment=ft.MainAxisAlignment.START,
            horizontal_alignment=ft.CrossAxisAlignment.STRETCH,
            spacing=0,
            scroll=ft.ScrollMode.HIDDEN,
        )

        self._left_panel = ft.Container(
            width=self.left_panel_width,
            bgcolor=BG_NAVY,
            padding=ft.padding.only(left=44, right=44, top=34, bottom=18),
            content=self._content_column,
        )

        self._hero_image = (
            ft.Image(
                src=self.image_path,
                fit="cover",
                width=self.image_width,
                height=self.content_height,
            )
            if has_side_image
            else None
        )
        self._hero_edge = ft.Container(
            width=6,
            height=self.content_height,
            bgcolor=BG_NAVY,
        )
        self._hero_overlay = ft.Container(
            width=self.image_width,
            height=self.content_height,
            gradient=ft.LinearGradient(
                begin=ft.Alignment(-1, 0),
                end=ft.Alignment(1, 0),
                colors=[
                    BG_NAVY,
                    ft.Colors.with_opacity(0.82, BG_NAVY),
                    ft.Colors.with_opacity(0.46, BG_NAVY),
                    ft.Colors.with_opacity(0.16, BG_NAVY),
                    ft.Colors.with_opacity(0.03, BG_NAVY),
                ],
                stops=[0.0, 0.02, 0.22, 0.56, 1.0],
            ),
        )
        self._right_stack = ft.Stack(
            [
                self._hero_image if self._hero_image is not None else ft.Container(expand=True),
                self._hero_edge,
                self._hero_overlay,
            ],
            width=self.image_width,
            height=self.content_height,
            clip_behavior=ft.ClipBehavior.HARD_EDGE,
        )
        self._right_panel = ft.Container(
            width=self.image_width,
            height=self.content_height,
            bgcolor=BG_NAVY,
            clip_behavior=ft.ClipBehavior.HARD_EDGE,
            content=self._right_stack,
        )

        self._root_row = ft.Row(
                [
                    self._left_panel,
                    self._right_panel,
                ],
                spacing=0,
                width=self.content_width,
                height=self.content_height,
        )
        self._root_container = ft.Container(
            width=self.content_width,
            height=self.content_height,
            bgcolor=BG_NAVY,
            padding=0,
            margin=0,
            content=self._root_row,
            clip_behavior=ft.ClipBehavior.HARD_EDGE,
        )
        page.add(self._root_container)

        self._queue_focus(self._user_field)
        self._start_update_check()
        page.run_task(self._poll_update_banner)
        self._fit_to_client_area()

    def mainloop(self):
        ft.app(target=self._main)


def main():
    app = LoginWindow()
    app.mainloop()


if __name__ == "__main__":
    main()
