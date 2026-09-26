import re
from io import BytesIO
from typing import Any

import requests
from docx import Document
from docx.enum.text import WD_BREAK
from docx.shared import Cm, Pt


SHARE_API_BASE = "https://chatgpt.com/backend-api/share/"


class ChatExportError(RuntimeError):
    """Понятная пользователю ошибка при получении или обработке shared-чата."""


def get_share_id(value: str) -> str:
    """Извлекает share ID из ссылки ChatGPT или принимает сам ID."""
    value = value.strip()
    if not value:
        raise ChatExportError("Вставьте shared-ссылку ChatGPT.")

    match = re.search(
        r"(?:https?://)?(?:www\.)?(?:chatgpt\.com|chat\.openai\.com)/share/([^/?#\s]+)",
        value,
        flags=re.IGNORECASE,
    )
    share_id = match.group(1) if match else value

    # Не даём случайно подставить произвольный URL/мусор в API-путь.
    if not re.fullmatch(r"[A-Za-z0-9_-]+", share_id):
        raise ChatExportError("Не удалось распознать shared-ссылку ChatGPT.")

    return share_id


def get_message_text(message: dict[str, Any]) -> str:
    content = message.get("content") or {}
    result: list[str] = []

    parts = content.get("parts", [])
    if isinstance(parts, list):
        for part in parts:
            if isinstance(part, str):
                if part.strip():
                    result.append(part)
            elif isinstance(part, dict):
                text = part.get("text")
                if isinstance(text, str) and text.strip():
                    result.append(text)

    direct_text = content.get("text")
    if not result and isinstance(direct_text, str) and direct_text.strip():
        result.append(direct_text)

    return "\n".join(result).strip()


def _message_is_exportable(message: dict[str, Any]) -> bool:
    author = message.get("author") or {}
    role = author.get("role")
    if role not in ("user", "assistant"):
        return False

    metadata = message.get("metadata") or {}
    if metadata.get("is_visually_hidden_from_conversation"):
        return False

    recipient = message.get("recipient")
    if role == "assistant" and recipient not in (None, "all"):
        return False

    return True


def get_messages_from_linear(data: dict[str, Any]) -> list[tuple[str, str]]:
    messages: list[tuple[str, str]] = []

    for item in data.get("linear_conversation", []):
        if not isinstance(item, dict):
            continue

        message = item.get("message")
        if not message and "author" in item:
            message = item

        if not isinstance(message, dict) or not _message_is_exportable(message):
            continue

        role = (message.get("author") or {}).get("role")
        text = get_message_text(message)
        if text:
            messages.append((role, text))

    return messages


def get_messages_from_mapping(data: dict[str, Any]) -> list[tuple[str, str]]:
    mapping = data.get("mapping") or {}
    if not mapping:
        return []

    current = data.get("current_node")
    nodes: list[dict[str, Any]] = []

    if current and current in mapping:
        while current:
            node = mapping.get(current)
            if not isinstance(node, dict):
                break
            nodes.append(node)
            current = node.get("parent")
        nodes.reverse()
    else:
        nodes = [node for node in mapping.values() if isinstance(node, dict)]

        def message_time(node: dict[str, Any]) -> float:
            message = node.get("message") or {}
            return message.get("create_time") or 0

        nodes.sort(key=message_time)

    messages: list[tuple[str, str]] = []
    for node in nodes:
        message = node.get("message")
        if not isinstance(message, dict) or not _message_is_exportable(message):
            continue

        role = (message.get("author") or {}).get("role")
        text = get_message_text(message)
        if text:
            messages.append((role, text))

    return messages


def fetch_shared_chat(shared_link: str) -> tuple[str, list[tuple[str, str]]]:
    """Загружает публичный shared-чат и возвращает название и сообщения."""
    share_id = get_share_id(shared_link)
    api_url = SHARE_API_BASE + share_id

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/153.0.0.0 Safari/537.36"
        ),
        "Accept": "application/json,text/plain,*/*",
    }

    try:
        response = requests.get(api_url, headers=headers, timeout=120)
    except requests.Timeout as exc:
        raise ChatExportError("ChatGPT слишком долго не отвечает. Попробуйте ещё раз.") from exc
    except requests.RequestException as exc:
        raise ChatExportError("Не удалось подключиться к ChatGPT.") from exc

    if response.status_code == 404:
        raise ChatExportError(
            "Shared-чат не найден. Проверьте ссылку и убедитесь, что общий доступ к чату включён."
        )

    if response.status_code == 403:
        raise ChatExportError(
            "ChatGPT отклонил запрос (403). Возможно, доступ к внутреннему share API изменился."
        )

    if response.status_code != 200:
        raise ChatExportError(f"ChatGPT вернул ошибку HTTP {response.status_code}.")

    try:
        data = response.json()
    except ValueError as exc:
        raise ChatExportError(
            "ChatGPT вернул ответ в неожиданном формате. Возможно, структура share API изменилась."
        ) from exc

    title = data.get("title") or "ChatGPT conversation"

    messages: list[tuple[str, str]] = []
    if data.get("linear_conversation"):
        messages = get_messages_from_linear(data)
    if not messages and data.get("mapping"):
        messages = get_messages_from_mapping(data)

    if not messages:
        raise ChatExportError(
            "Чат загрузился, но текстовые сообщения не найдены. Возможно, структура share API изменилась."
        )

    return title, messages


def safe_filename(name: str) -> str:
    name = re.sub(r'[<>:"/\\|?*]', "_", name).strip(" .")
    return name[:150] or "ChatGPT conversation"


def build_docx(title: str, messages: list[tuple[str, str]]) -> bytes:
    """Собирает Word-документ в памяти и возвращает его как bytes."""
    doc = Document()

    section = doc.sections[0]
    section.top_margin = Cm(1.8)
    section.bottom_margin = Cm(1.8)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)

    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(11)

    title_paragraph = doc.add_heading(title, level=0)
    title_paragraph.paragraph_format.space_after = Pt(18)

    for index, (role, text) in enumerate(messages):
        speaker = "Вы" if role == "user" else "ChatGPT"

        speaker_paragraph = doc.add_paragraph()
        speaker_paragraph.paragraph_format.space_before = Pt(6)
        speaker_paragraph.paragraph_format.space_after = Pt(4)
        speaker_run = speaker_paragraph.add_run(speaker)
        speaker_run.bold = True
        speaker_run.font.size = Pt(12)

        body = doc.add_paragraph(text)
        body.paragraph_format.space_after = Pt(10)
        body.paragraph_format.line_spacing = 1.08

        if index != len(messages) - 1:
            separator = doc.add_paragraph()
            separator.paragraph_format.space_after = Pt(3)
            separator.add_run("─" * 48)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def export_shared_chat(shared_link: str) -> tuple[str, str, int, bytes]:
    """Полный сценарий: ссылка -> чат -> DOCX."""
    title, messages = fetch_shared_chat(shared_link)
    filename = safe_filename(title) + ".docx"
    docx_bytes = build_docx(title, messages)
    return title, filename, len(messages), docx_bytes
