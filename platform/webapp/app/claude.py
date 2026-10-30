# -*- coding: utf-8 -*-
"""Claude API 호출 — 샘플 서류 + 입력값으로 '샘플과 비슷한 양식'의 결과물을 스트리밍 생성한다.
- 모델: config.MODEL (기본 claude-opus-5-5), effort: config.EFFORT
- 안전 분류기가 요청을 거절하면 서버 측 fallback(`fallbacks: "default"`)이 다른 모델로 다시 실행한다.
- 시장조사 모듈은 웹 검색(server tool)을 켤 수 있다."""
import base64
from pathlib import Path
from typing import Iterator

import anthropic

from . import config
from .extract import extract

SYSTEM = """당신은 부동산 자산운용사(리츠·부동산펀드) 투자팀 운용역을 돕는 실무 보조입니다. 한국어로 작성합니다.

작성 원칙
1. 샘플 서류가 주어지면 그 양식(목차·조항 번호 체계·표 구성·문체·머리글/서명 형식)을 최대한 그대로 따르고, 내용만 이번 입력값으로 바꿉니다.
2. 수치에는 출처(파일명·자료명·기준시점)를 붙입니다. 입력에 없는 값은 지어내지 말고 [  ](미정), [추정], [확인 필요], [데이터 대기]로 표시합니다.
3. 수익률·배당 등에 '보장', '확정' 같은 확약성 표현을 쓰지 않습니다.
4. 법률·세무 판단은 내부 검토이며 외부 법무법인·회계법인 확인이 필요하다고 마지막에 한 줄 적습니다(해당하는 경우).
5. 금액은 원 단위 천 단위 쉼표 또는 억원 표기를 일관되게 쓰고, 날짜는 YYYY-MM-DD 또는 YYYY년 M월 D일로 씁니다.
6. 출력은 Markdown으로 합니다. 제목은 #, 표는 | 구분 표, 목록은 - 또는 1. 을 씁니다. 결과물 본문만 출력하고 앞뒤 인사말은 쓰지 않습니다.
7. 마지막에 '## 확인 필요 사항' 절을 두어 미정·추정 항목과 사람이 확인해야 할 점을 정리합니다."""

FALLBACK_BETA = "server-side-fallback-2026-07-01"
WEB_SEARCH = {"type": "web_search_20260209", "name": "web_search", "max_uses": 8}
MAX_PDF_BYTES = 30 * 1024 * 1024


def build_content(module: dict, subtype: dict, fields: dict, samples: list[dict], extra: str, include_text: str) -> tuple[list, list]:
    """samples: [{'name': 표시명, 'path': Path}] → (content blocks, warnings)"""
    blocks, warnings = [], []
    for i, s in enumerate(samples, 1):
        path = Path(s["path"])
        info = extract(path)
        if info["needs_pdf"] and path.stat().st_size <= MAX_PDF_BYTES:
            blocks.append({"type": "document", "title": s["name"],
                           "source": {"type": "base64", "media_type": "application/pdf",
                                      "data": base64.standard_b64encode(path.read_bytes()).decode()}})
            continue
        if info["needs_pdf"]:
            warnings.append(f"{s['name']}: 스캔 PDF가 30MB를 넘어 글자를 읽지 못했습니다.")
        if info["truncated"]:
            warnings.append(f"{s['name']}: 분량이 많아 앞부분 {config.MAX_SAMPLE_CHARS:,}자만 참고했습니다.")
        blocks.append({"type": "text", "text": f"<sample index=\"{i}\" name=\"{s['name']}\">\n{info['text']}\n</sample>"})

    lines = [f"[업무 단계] {module['name']} — {subtype['name']}", f"[작성 지침] {subtype['instructions']}", "", "[입력값]"]
    labels = {f["key"]: f["label"] for f in subtype.get("fields", [])}
    for k, v in fields.items():
        if str(v).strip():
            lines.append(f"- {labels.get(k, k)}: {v}")
    if include_text:
        lines += ["", include_text]
    if extra.strip():
        lines += ["", f"[추가 요청] {extra}"]
    if samples:
        lines += ["", "위 <sample> 서류(또는 첨부 PDF)의 양식을 기준으로 작성하세요."]
    else:
        lines += ["", "샘플 서류가 없으므로 실무 표준 양식으로 작성하세요."]
    blocks.append({"type": "text", "text": "\n".join(lines)})
    return blocks, warnings


def stream_generate(content: list, web_search: bool = False) -> Iterator[str]:
    """텍스트 조각을 차례로 내보낸다. 오류·거절은 '[안내]' 문장으로 내보낸다."""
    try:
        client = anthropic.Anthropic()
    except Exception as e:
        yield f"\n\n[안내] Claude API 인증 정보가 없습니다. platform/webapp/.env 에 ANTHROPIC_API_KEY 를 넣어 주세요. ({e})"
        return
    messages = [{"role": "user", "content": content}]
    kwargs = dict(model=config.MODEL, max_tokens=64000, system=SYSTEM, messages=messages,
                  betas=[FALLBACK_BETA], fallbacks="default",
                  output_config={"effort": config.EFFORT}, cache_control={"type": "ephemeral"})
    if web_search:
        kwargs["tools"] = [WEB_SEARCH]
    try:
        for _ in range(4):  # 웹 검색이 길어져 pause_turn 이 오면 이어서 실행
            with client.beta.messages.stream(**kwargs) as stream:
                for text in stream.text_stream:
                    yield text
                final = stream.get_final_message()
            if final.stop_reason == "pause_turn":
                messages.append({"role": "assistant", "content": [b.to_dict() for b in final.content]})
                continue
            if final.stop_reason == "refusal":
                yield "\n\n[안내] 안전 정책으로 이 요청은 생성되지 않았습니다. 입력 내용을 바꿔 다시 시도해 주세요."
            elif final.stop_reason == "max_tokens":
                yield "\n\n[안내] 출력 길이 한도에 도달해 결과가 잘렸습니다. 범위를 나눠 다시 생성해 주세요."
            break
    except anthropic.AuthenticationError:
        yield "\n\n[안내] API 키가 올바르지 않습니다. .env 의 ANTHROPIC_API_KEY 를 확인해 주세요."
    except anthropic.PermissionDeniedError as e:
        yield f"\n\n[안내] 이 API 키로 사용할 수 없는 기능·모델입니다: {e.message}"
    except anthropic.RateLimitError:
        yield "\n\n[안내] 요청 한도를 초과했습니다. 1~2분 뒤 다시 시도해 주세요."
    except anthropic.BadRequestError as e:
        yield f"\n\n[안내] 요청 형식 오류: {e.message}"
    except anthropic.APIStatusError as e:
        yield f"\n\n[안내] Claude API 오류({e.status_code}): {e.message}"
    except anthropic.APIConnectionError:
        yield "\n\n[안내] Claude API에 연결하지 못했습니다. 인터넷 연결을 확인해 주세요."
    except TypeError as e:  # 인증 정보가 하나도 없을 때 SDK가 요청 직전에 내는 오류
        if "authentication" in str(e):
            yield "\n\n[안내] Claude API 인증 정보가 없습니다. platform/webapp/.env 에 ANTHROPIC_API_KEY 를 넣고 다시 실행해 주세요."
        else:
            yield f"\n\n[안내] 처리 중 오류: {e}"
    except Exception as e:  # 예상하지 못한 오류도 화면에 알리고 스트림은 정상 종료
        yield f"\n\n[안내] 처리 중 오류: {e}"
