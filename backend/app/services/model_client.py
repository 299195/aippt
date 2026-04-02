from __future__ import annotations

import base64
import http.client
import json
import mimetypes
import shutil
import socket
import ssl
import subprocess
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterator
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

from app.config import settings


def _extract_json_from_text(text: str) -> dict[str, Any]:
    payload = (text or "").strip()
    if not payload:
        raise ValueError("empty model response")

    if payload.startswith("```"):
        lines = payload.splitlines()
        if lines and lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].strip() == "```":
            lines = lines[:-1]
        payload = "\n".join(lines).strip()

    candidates: list[str] = [payload]

    start_obj = payload.find("{")
    end_obj = payload.rfind("}")
    if start_obj >= 0 and end_obj > start_obj:
        candidates.append(payload[start_obj : end_obj + 1])

    start_arr = payload.find("[")
    end_arr = payload.rfind("]")
    if start_arr >= 0 and end_arr > start_arr:
        candidates.append(payload[start_arr : end_arr + 1])

    last_err: Exception | None = None
    for candidate in candidates:
        try:
            parsed = json.loads(candidate)
            if isinstance(parsed, dict):
                return parsed
            raise ValueError("json response is not an object")
        except Exception as exc:  # noqa: PERF203
            last_err = exc

    raise ValueError(f"cannot parse json object: {last_err}")


def _guess_mime_type(path: Path) -> str:
    mime, _ = mimetypes.guess_type(str(path))
    if mime:
        return mime
    return "image/png"


@dataclass
class ModelClient:
    provider: str = settings.model_provider
    base_url: str = settings.model_base_url
    api_key: str = settings.model_api_key
    model: str = settings.model_name
    endpoint_id: str = settings.model_endpoint_id

    def _target_model(self) -> str:
        # Volcengine Ark endpoint mode uses endpoint_id as the OpenAI-compatible model field.
        return str(self.endpoint_id or self.model or "").strip()

    def enabled(self) -> bool:
        return bool(self.base_url and self.api_key and self._target_model())

    def _post(self, payload: dict[str, Any]) -> dict[str, Any]:
        if not self.enabled():
            raise RuntimeError(f"{self.provider} model config incomplete")

        url = self.base_url.rstrip("/") + settings.model_chat_path
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        ssl_context = ssl.create_default_context()
        last_error: Exception | None = None

        for attempt in range(1, 4):
            req = Request(url=url, data=body, method="POST")
            req.add_header("Content-Type", "application/json")
            req.add_header("Authorization", f"Bearer {self.api_key}")
            try:
                with urlopen(req, timeout=settings.request_timeout_sec, context=ssl_context) as resp:
                    return json.loads(resp.read().decode("utf-8"))
            except (TimeoutError, socket.timeout) as exc:
                last_error = exc
                if attempt < 3:
                    time.sleep(0.6 * attempt)
                    continue
                raise RuntimeError(
                    f"{self.provider} request timeout after {settings.request_timeout_sec}s: {exc}"
                ) from exc
            except HTTPError as exc:
                detail = exc.read().decode("utf-8", errors="ignore")
                transient = exc.code in {408, 409, 425, 429} or exc.code >= 500
                if transient and attempt < 3:
                    time.sleep(0.6 * attempt)
                    continue
                raise RuntimeError(f"{self.provider} HTTPError: {exc.code} {detail}") from exc
            except URLError as exc:
                last_error = exc
                if attempt < 3:
                    time.sleep(0.6 * attempt)
                    continue
                raise RuntimeError(f"{self.provider} URLError: {exc}") from exc
            except (http.client.RemoteDisconnected, ConnectionResetError) as exc:
                last_error = exc
                if attempt < 3:
                    time.sleep(0.6 * attempt)
                    continue
                raise RuntimeError(
                    f"{self.provider} connection closed by remote host: {exc}. "
                    f"target={url}"
                ) from exc

        raise RuntimeError(f"{self.provider} request failed: {last_error}")

    @staticmethod
    def _extract_content(resp: dict[str, Any]) -> str:
        choices = resp.get("choices")
        if not isinstance(choices, list) or not choices:
            raise RuntimeError("model response missing choices")

        message = choices[0].get("message", {})
        content = message.get("content", "")

        if isinstance(content, str):
            return content.strip()

        if isinstance(content, list):
            parts: list[str] = []
            for block in content:
                if not isinstance(block, dict):
                    continue
                text = block.get("text")
                if isinstance(text, str):
                    parts.append(text)
            return "\n".join(parts).strip()

        return str(content).strip()

    @staticmethod
    def _extract_delta_text(event: dict[str, Any]) -> str:
        choices = event.get("choices")
        if not isinstance(choices, list) or not choices:
            return ""

        delta = choices[0].get("delta", {})
        content = delta.get("content", "")
        if isinstance(content, str):
            return content

        if isinstance(content, list):
            parts: list[str] = []
            for block in content:
                if not isinstance(block, dict):
                    continue
                text = block.get("text")
                if isinstance(text, str):
                    parts.append(text)
            return "".join(parts)

        return ""

    @staticmethod
    def _image_to_data_url(image_path: str | Path) -> str:
        path = Path(image_path)
        if not path.exists() or not path.is_file():
            raise FileNotFoundError(f"image not found: {path}")

        raw = path.read_bytes()
        b64 = base64.b64encode(raw).decode("ascii")
        mime = _guess_mime_type(path)
        return f"data:{mime};base64,{b64}"

    def chat_text(
        self,
        system_prompt: str,
        user_prompt: str,
        temperature: float = 0.3,
        response_format: dict[str, Any] | None = None,
    ) -> str:
        payload: dict[str, Any] = {
            "model": self._target_model(),
            "temperature": temperature,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
        }
        if response_format:
            payload["response_format"] = response_format

        data = self._post(payload)
        content = self._extract_content(data)
        if not content:
            raise RuntimeError(f"{self.provider} returned empty content")
        return content

    def _chat_stream_payload(
        self,
        system_prompt: str,
        user_prompt: str,
        temperature: float,
    ) -> dict[str, Any]:
        return {
            "model": self._target_model(),
            "temperature": temperature,
            "stream": True,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
        }

    def _chat_text_stream_with_curl(
        self,
        *,
        url: str,
        payload: dict[str, Any],
        stream_timeout: int,
    ) -> Iterator[str]:
        curl_bin = shutil.which("curl") or shutil.which("curl.exe")
        if not curl_bin:
            raise RuntimeError("curl executable not found")

        body = json.dumps(payload, ensure_ascii=False)
        cmd = [
            curl_bin,
            "-sS",
            "-N",
            "--connect-timeout",
            str(max(5, min(stream_timeout, 30))),
            "--max-time",
            str(max(15, stream_timeout)),
            "-X",
            "POST",
            url,
            "-H",
            "Content-Type: application/json",
            "-H",
            f"Authorization: Bearer {self.api_key}",
            "--data-binary",
            body,
        ]

        try:
            proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding="utf-8",
                errors="ignore",
                bufsize=1,
            )
        except Exception as exc:  # noqa: PERF203
            raise RuntimeError(f"{self.provider} failed to start curl: {exc}") from exc

        lines: list[str] = []
        saw_sse = False
        try:
            assert proc.stdout is not None
            for raw_line in proc.stdout:
                line = raw_line.strip()
                if not line:
                    continue
                lines.append(line)

                if not line.startswith("data:"):
                    continue

                saw_sse = True
                data = line[5:].strip()
                if data == "[DONE]":
                    break
                try:
                    event = json.loads(data)
                except Exception:
                    continue

                chunk = self._extract_delta_text(event)
                if chunk:
                    yield chunk
        finally:
            stderr_text = ""
            if proc.stderr is not None:
                try:
                    stderr_text = proc.stderr.read().strip()
                except Exception:
                    stderr_text = ""
            return_code = proc.wait()
            if return_code != 0:
                detail = stderr_text or f"curl exit code {return_code}"
                raise RuntimeError(f"{self.provider} curl stream failed: {detail}")

        # Some providers may ignore stream=true and return a normal JSON body.
        if not saw_sse and lines:
            joined = "\n".join(lines)
            try:
                data = json.loads(joined)
                text = self._extract_content(data)
                if text:
                    yield text
            except Exception:
                pass

    def _chat_text_stream_with_urllib(
        self,
        *,
        url: str,
        payload: dict[str, Any],
        stream_timeout: int,
    ) -> Iterator[str]:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        req = Request(url=url, data=body, method="POST")
        req.add_header("Content-Type", "application/json")
        req.add_header("Authorization", f"Bearer {self.api_key}")

        ssl_context = ssl.create_default_context()
        with urlopen(req, timeout=stream_timeout, context=ssl_context) as resp:
            lines: list[str] = []
            saw_sse = False

            for raw_line in resp:
                line = raw_line.decode("utf-8", errors="ignore").strip()
                if not line:
                    continue
                lines.append(line)

                if line.startswith("data:"):
                    saw_sse = True
                    data = line[5:].strip()
                    if data == "[DONE]":
                        break
                    try:
                        event = json.loads(data)
                    except Exception:
                        continue

                    chunk = self._extract_delta_text(event)
                    if chunk:
                        yield chunk
                    continue

            # Some providers may ignore stream=true and return a normal JSON body.
            if not saw_sse and lines:
                joined = "\n".join(lines)
                try:
                    data = json.loads(joined)
                    text = self._extract_content(data)
                    if text:
                        yield text
                except Exception:
                    pass

    def chat_text_stream(
        self,
        system_prompt: str,
        user_prompt: str,
        temperature: float = 0.3,
    ) -> Iterator[str]:
        if not self.enabled():
            raise RuntimeError(f"{self.provider} model config incomplete")

        url = self.base_url.rstrip("/") + settings.model_chat_path
        payload = self._chat_stream_payload(system_prompt, user_prompt, temperature)
        stream_timeout = max(int(settings.stream_timeout_sec), int(settings.request_timeout_sec))

        try:
            # third-party-compatible path: prefer curl streaming, fallback to urllib.
            yield from self._chat_text_stream_with_curl(
                url=url,
                payload=payload,
                stream_timeout=stream_timeout,
            )
        except HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="ignore")
            raise RuntimeError(f"{self.provider} HTTPError: {exc.code} {detail}") from exc
        except (TimeoutError, socket.timeout) as exc:
            raise RuntimeError(
                f"{self.provider} stream timeout after {stream_timeout}s: {exc}. "
                "Increase MODEL_STREAM_TIMEOUT if needed."
            ) from exc
        except URLError as exc:
            raise RuntimeError(f"{self.provider} URLError: {exc}") from exc
        except (http.client.RemoteDisconnected, ConnectionResetError) as exc:
            raise RuntimeError(
                f"{self.provider} stream closed by remote host: {exc}. "
                f"target={url}"
            ) from exc
        except Exception:
            # Curl path may be unavailable on some environments.
            try:
                yield from self._chat_text_stream_with_urllib(
                    url=url,
                    payload=payload,
                    stream_timeout=stream_timeout,
                )
            except HTTPError as exc:
                detail = exc.read().decode("utf-8", errors="ignore")
                raise RuntimeError(f"{self.provider} HTTPError: {exc.code} {detail}") from exc
            except (TimeoutError, socket.timeout) as exc:
                raise RuntimeError(
                    f"{self.provider} stream timeout after {stream_timeout}s: {exc}. "
                    "Increase MODEL_STREAM_TIMEOUT if needed."
                ) from exc
            except URLError as exc:
                raise RuntimeError(f"{self.provider} URLError: {exc}") from exc
            except (http.client.RemoteDisconnected, ConnectionResetError) as exc:
                raise RuntimeError(
                    f"{self.provider} stream closed by remote host: {exc}. "
                    f"target={url}"
                ) from exc

    def chat_with_image_text(
        self,
        system_prompt: str,
        user_prompt: str,
        image_path: str | Path,
        temperature: float = 0.2,
    ) -> str:
        image_url = self._image_to_data_url(image_path)

        payload: dict[str, Any] = {
            "model": self._target_model(),
            "temperature": temperature,
            "messages": [
                {"role": "system", "content": system_prompt},
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": user_prompt},
                        {"type": "image_url", "image_url": {"url": image_url}},
                    ],
                },
            ],
        }

        data = self._post(payload)
        content = self._extract_content(data)
        if not content:
            raise RuntimeError(f"{self.provider} returned empty content")
        return content

    def chat_json(self, system_prompt: str, user_prompt: str, temperature: float = 0.3) -> Dict[str, Any]:
        content = self.chat_text(
            system_prompt=system_prompt,
            user_prompt=user_prompt,
            temperature=temperature,
            response_format={"type": "json_object"},
        )
        return _extract_json_from_text(content)
