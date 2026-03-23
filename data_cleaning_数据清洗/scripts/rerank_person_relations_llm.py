from __future__ import annotations

import json
import math
import os
import re
import subprocess
import sys
import time
from pathlib import Path
from typing import Any

import pandas as pd
import requests
from openai import OpenAI


PROJECT_ROOT = Path(__file__).resolve().parents[2]
KB_DIR = PROJECT_ROOT / "output_输出结果" / "kb_data_知识库数据"
REPORT_DIR = PROJECT_ROOT / "output_输出结果" / "reports_报告"
APP_DIR = PROJECT_ROOT / "knowledge_base_知识库构建" / "app"

PERSONS_PATH = KB_DIR / "persons.csv"
RELATIONS_PATH = KB_DIR / "person_relations.csv"
SOURCES_PATH = KB_DIR / "sources.csv"
REPORT_PATH = REPORT_DIR / "llm_relation_rerank_report.md"

LABELS = [
    "同属组织",
    "通信",
    "合作",
    "交往",
    "论战",
    "共同活动",
    "亲属",
    "师生",
    "悼念/纪念关联",
    "仅共现",
    "证据不足",
    "冲突未解",
]
HIDDEN_LABELS = {"仅共现", "证据不足", "冲突未解"}
CONF_MAP = {"low": 0.40, "medium": 0.70, "high": 0.95}
EVIDENCE_STRENGTH_LEVELS = {"weak": 1, "moderate": 2, "strong": 3}
DEFAULT_BATCH_SIZE = 10
DEFAULT_SLEEP_SECONDS = 0.35
DEFAULT_MODEL_CANDIDATES = [
    "claude-sonnet-4-5-20251101",
    "claude-haiku-4-5-20251101",
    "claude-opus-4-5-20251101",
    "claude-3-7-sonnet-20250219",
    "claude-3-5-sonnet-20241022",
]
DEFAULT_OPENAI_MODELS = ["gpt-4o", "gpt-4.1", "gpt-4.1-mini", "gpt-4o-mini", "o4-mini"]


def text(value: Any) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def confidence_to_float(value: Any) -> float:
    raw = text(value).lower()
    if raw in CONF_MAP:
        return CONF_MAP[raw]
    try:
        return float(raw)
    except ValueError:
        return 0.0


def build_pair_key(person_a: str, person_b: str) -> str:
    return "__".join(sorted([text(person_a), text(person_b)]))


def split_ids(value: str) -> list[str]:
    if not value:
        return []
    return [item.strip() for item in str(value).split(";") if item.strip()]


def candidate_base_urls(base_url: str) -> list[str]:
    base = base_url.rstrip("/")
    candidates = [base]
    marker = "https://code.newcli.com/claude"
    if base.startswith(marker):
        for variant in [
            marker,
            f"{marker}/super",
            f"{marker}/aws",
            f"{marker}/ultra",
            f"{marker}/droid",
        ]:
            if variant not in candidates:
                candidates.append(variant)
    return candidates


def normalize_endpoint(base_url: str) -> list[str]:
    candidates: list[str] = []
    for base in candidate_base_urls(base_url):
        if base.endswith("/v1/messages") or base.endswith("/messages"):
            if base not in candidates:
                candidates.append(base)
            continue
        for endpoint in [f"{base}/v1/messages", f"{base}/messages"]:
            if endpoint not in candidates:
                candidates.append(endpoint)
    return candidates


def get_models() -> list[str]:
    env_model = text(os.getenv("ANTHROPIC_MODEL"))
    if env_model:
        return [item.strip() for item in env_model.split(",") if item.strip()]
    return DEFAULT_MODEL_CANDIDATES


def get_openai_models() -> list[str]:
    env_model = text(os.getenv("OPENAI_MODEL") or os.getenv("RELATION_RERANK_MODEL"))
    strict_model = text(os.getenv("RELATION_RERANK_STRICT_MODEL")).lower() in {"1", "true", "yes", "on"}
    models: list[str] = []
    if env_model:
        models.extend([item.strip() for item in env_model.split(",") if item.strip()])
    if strict_model:
        return models
    for model in DEFAULT_OPENAI_MODELS:
        if model not in models:
            models.append(model)
    return models


def load_frames() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    persons = pd.read_csv(PERSONS_PATH).fillna("")
    relations = pd.read_csv(RELATIONS_PATH).fillna("")
    sources = pd.read_csv(SOURCES_PATH).fillna("")
    return persons, relations, sources


def ensure_columns(relations: pd.DataFrame) -> pd.DataFrame:
    raw_default = relations["original_relation_type"] if "original_relation_type" in relations.columns else relations.get("standard_relation_type", "")
    current_default = relations["standard_relation_type"] if "standard_relation_type" in relations.columns else raw_default
    defaults = {
        "raw_relation_type": raw_default,
        "llm_suggested_relation_type": "",
        "final_relation_type": current_default,
        "llm_reason": "",
        "llm_confidence": "",
        "display_status": "formal",
    }
    updated = relations.copy()
    for column, default_value in defaults.items():
        if column not in updated.columns:
            updated[column] = default_value
    if "original_relation_type" in updated.columns:
        updated["raw_relation_type"] = updated["raw_relation_type"].where(
            updated["raw_relation_type"].astype(str).str.strip() != "",
            updated["original_relation_type"],
        )
    if "standard_relation_type" in updated.columns:
        updated["final_relation_type"] = updated["final_relation_type"].where(
            updated["final_relation_type"].astype(str).str.strip() != "",
            updated["standard_relation_type"],
        )
    updated["final_relation_type"] = updated["final_relation_type"].where(
        updated["final_relation_type"].astype(str).str.strip() != "",
        updated["raw_relation_type"],
    )
    updated["display_status"] = updated["display_status"].replace("", "formal")
    return updated


def target_mask(relations: pd.DataFrame) -> pd.Series:
    if "final_relation_type" in relations.columns:
        current_relation = relations["final_relation_type"].astype(str)
    elif "standard_relation_type" in relations.columns:
        current_relation = relations["standard_relation_type"].astype(str)
    else:
        current_relation = relations["raw_relation_type"].astype(str)
    manual_review = relations["needs_manual_review"].astype(str).str.lower().eq("yes")
    low_conf = relations["confidence"].apply(confidence_to_float) < 0.75
    return current_relation.str.startswith("待核") & manual_review & low_conf


def rerank_scope_mask(relations: pd.DataFrame) -> pd.Series:
    if "final_relation_type" in relations.columns:
        current_relation = relations["standard_relation_type"].astype(str) if "standard_relation_type" in relations.columns else relations["final_relation_type"].astype(str)
    elif "standard_relation_type" in relations.columns:
        current_relation = relations["standard_relation_type"].astype(str)
    else:
        current_relation = relations["raw_relation_type"].astype(str)
    manual_review = relations["needs_manual_review"].astype(str).str.lower().eq("yes")
    low_conf = relations["confidence"].apply(confidence_to_float) < 0.75
    return current_relation.str.startswith("待核") & manual_review & low_conf


def source_index(sources: pd.DataFrame) -> dict[str, dict[str, str]]:
    return {
        text(row["source_id"]): {key: text(value) for key, value in row.items()}
        for _, row in sources.iterrows()
        if text(row["source_id"])
    }


def build_local_evidence(source_ids: str, source_lookup: dict[str, dict[str, str]]) -> tuple[list[str], list[str]]:
    local_evidence: list[str] = []
    urls: list[str] = []
    seen_local: set[str] = set()
    seen_url: set[str] = set()
    for source_id in split_ids(source_ids):
        row = source_lookup.get(source_id, {})
        citation = row.get("citation", "")
        source_path = row.get("source_path", "")
        title = row.get("title", "")
        source_url = row.get("source_url", "")
        if source_path or citation:
            item = " | ".join([part for part in [title, citation, source_path] if part])
            if item and item not in seen_local:
                seen_local.add(item)
                local_evidence.append(item)
        if source_url and source_url not in seen_url:
            seen_url.add(source_url)
            urls.append(source_url)
    return local_evidence[:8], urls[:6]


def infer_evidence_strength(record: dict[str, Any]) -> str:
    evidence_count = len(record["evidence_texts"])
    local_count = len(record["local_evidence"])
    url_count = len(record["source_url"])
    max_weight = max(record["weights"]) if record["weights"] else 0
    if (evidence_count >= 2 and local_count >= 2) or (evidence_count >= 2 and url_count >= 1) or max_weight >= 4:
        return "strong"
    if evidence_count >= 1 and (local_count >= 1 or max_weight >= 2):
        return "moderate"
    return "weak"


def aggregate_pairs(
    relations: pd.DataFrame,
    persons: pd.DataFrame,
    sources: pd.DataFrame,
) -> tuple[list[dict[str, Any]], dict[str, str]]:
    name_map = dict(zip(persons["person_id"], persons["standard_name"]))
    role_map = dict(zip(persons["person_id"], persons["role"]))
    source_lookup = source_index(sources)
    targets = relations[target_mask(relations)].copy()
    targets["pair_key"] = targets.apply(lambda row: build_pair_key(row["source_person_id"], row["target_person_id"]), axis=1)

    pair_payloads: list[dict[str, Any]] = []
    pair_reason_map: dict[str, str] = {}

    for pair_key, group in targets.groupby("pair_key", sort=True):
        source_person_id, target_person_id = pair_key.split("__", 1)
        raw_types = []
        evidence_texts = []
        source_ids = []
        weights = []
        for _, row in group.iterrows():
            raw_types.append(text(row["raw_relation_type"]))
            context = text(row["context"])
            if context:
                evidence_texts.append(context[:420])
            source_ids.extend(split_ids(text(row["source_ids"])))
            try:
                weights.append(float(text(row["weight"]) or "0"))
            except ValueError:
                weights.append(0.0)

        local_evidence, source_urls = build_local_evidence(";".join(source_ids), source_lookup)
        payload = {
            "pair_key": pair_key,
            "source_person_id": source_person_id,
            "target_person_id": target_person_id,
            "source_person": text(name_map.get(source_person_id, source_person_id)),
            "target_person": text(name_map.get(target_person_id, target_person_id)),
            "source_person_role": text(role_map.get(source_person_id, "")),
            "target_person_role": text(role_map.get(target_person_id, "")),
            "original_relation_type": sorted({item for item in raw_types if item}),
            "evidence_texts": evidence_texts[:4],
            "local_evidence": local_evidence,
            "source_url": source_urls,
            "candidate_relation_labels": LABELS,
            "weights": weights,
            "group_size": len(group),
            "heuristic_evidence_strength": "",
        }
        payload["heuristic_evidence_strength"] = infer_evidence_strength(payload)
        pair_payloads.append(payload)
        pair_reason_map[pair_key] = " | ".join(sorted({text(row["correction_reason"]) for _, row in group.iterrows() if text(row["correction_reason"])}))[:500]

    return pair_payloads, pair_reason_map


def strip_json_wrapper(text_value: str) -> str:
    cleaned = text_value.strip()
    cleaned = re.sub(r"^```json\s*", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"^```\s*", "", cleaned)
    cleaned = re.sub(r"\s*```$", "", cleaned)
    return cleaned.strip()


def call_anthropic_batch(pair_batch: list[dict[str, Any]], base_url: str, auth_token: str, model_candidates: list[str]) -> dict[str, Any]:
    if not base_url or not auth_token:
        raise RuntimeError("缺少 ANTHROPIC_BASE_URL / ANTHROPIC_AUTH_TOKEN。")

    system_prompt = (
        "你是左联人物关系重判器。"
        "你只能在给定标签集合中选一个标签，不允许发明新标签。"
        "禁止输出“组织隶属”，如果证据仅能说明共同属于某组织，必须输出“同属组织”。"
        "如果只有并列名单或弱共现，输出“仅共现”或“证据不足”。"
        "如果证据互相矛盾无法消解，输出“冲突未解”。"
        "请仅输出 JSON。"
    )
    user_payload = {
        "task": "对每个人物对给出唯一标签、置信度、证据强度与简要理由。",
        "candidate_relation_labels": LABELS,
        "pairs": pair_batch,
        "output_schema": {
            "results": [
                {
                    "pair_key": "string",
                    "llm_suggested_relation_type": f"enum({', '.join(LABELS)})",
                    "llm_confidence": "0-1 float",
                    "evidence_strength": "enum(strong, moderate, weak)",
                    "llm_reason": "short string",
                }
            ]
        },
    }

    headers = {
        "content-type": "application/json",
        "x-api-key": auth_token,
        "Authorization": f"Bearer {auth_token}",
        "anthropic-version": "2023-06-01",
    }

    errors: list[str] = []
    for endpoint in normalize_endpoint(base_url):
        for model in model_candidates:
            payload = {
                "model": model,
                "max_tokens": 3500,
                "temperature": 0,
                "system": system_prompt,
                "messages": [
                    {
                        "role": "user",
                        "content": json.dumps(user_payload, ensure_ascii=False),
                    }
                ],
            }
            try:
                response = requests.post(endpoint, headers=headers, json=payload, timeout=120)
            except requests.RequestException as exc:
                errors.append(f"{endpoint} | {model} | request_error={exc}")
                continue

            if response.status_code >= 400:
                errors.append(f"{endpoint} | {model} | status={response.status_code} | body={response.text[:400]}")
                continue

            body = response.json()
            content_items = body.get("content", [])
            texts = [item.get("text", "") for item in content_items if item.get("type") == "text"]
            if not texts:
                errors.append(f"{endpoint} | {model} | empty_text_response")
                continue
            raw_text = strip_json_wrapper("\n".join(texts))
            try:
                parsed = json.loads(raw_text)
            except json.JSONDecodeError as exc:
                errors.append(f"{endpoint} | {model} | invalid_json={exc} | raw={raw_text[:400]}")
                continue
            parsed["_endpoint"] = endpoint
            parsed["_model"] = model
            return parsed

    raise RuntimeError(" ; ".join(errors[-6:]) if errors else "anthropic batch call failed")


def call_openai_batch(pair_batch: list[dict[str, Any]], base_url: str, api_key: str, model_candidates: list[str]) -> dict[str, Any]:
    if not base_url or not api_key:
        raise RuntimeError("缺少 OPENAI_BASE_URL / OPENAI_API_KEY。")
    if not model_candidates:
        raise RuntimeError("缺少 OPENAI_MODEL / RELATION_RERANK_MODEL。")

    system_prompt = (
        "你是左联人物关系重判器。"
        "你只能从给定标签集合中选择一个标签，不允许发明新标签。"
        "禁止输出“组织隶属”，如果证据只支持共同属于某组织，必须输出“同属组织”。"
        "如果只是并列名单或弱共现，输出“仅共现”或“证据不足”。"
        "如果证据冲突无法消解，输出“冲突未解”。"
        "只输出 JSON，不要输出 markdown。"
    )
    user_payload = {
        "task": "对每个人物对给出唯一标签、置信度、证据强度与简要理由。",
        "candidate_relation_labels": LABELS,
        "pairs": pair_batch,
        "output_schema": {
            "results": [
                {
                    "pair_key": "string",
                    "llm_suggested_relation_type": f"enum({', '.join(LABELS)})",
                    "llm_confidence": "0-1 float",
                    "evidence_strength": "enum(strong, moderate, weak)",
                    "llm_reason": "short string",
                }
            ]
        },
    }

    client = OpenAI(api_key=api_key, base_url=base_url)
    errors: list[str] = []
    for model in model_candidates:
        try:
            response = client.chat.completions.create(
                model=model,
                temperature=0,
                response_format={"type": "json_object"},
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": json.dumps(user_payload, ensure_ascii=False)},
                ],
            )
            content = response.choices[0].message.content or ""
            parsed = json.loads(strip_json_wrapper(content))
            parsed["_model"] = model
            parsed["_endpoint"] = base_url
            return parsed
        except Exception as exc:
            errors.append(f"{model} | json_object | {exc}")
        try:
            response = client.chat.completions.create(
                model=model,
                temperature=0,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": json.dumps(user_payload, ensure_ascii=False)},
                ],
            )
            content = response.choices[0].message.content or ""
            parsed = json.loads(strip_json_wrapper(content))
            parsed["_model"] = model
            parsed["_endpoint"] = base_url
            return parsed
        except Exception as exc:
            errors.append(f"{model} | plain_json | {exc}")

    if not errors:
        raise RuntimeError("openai batch call failed")
    if len(errors) <= 8:
        raise RuntimeError(" ; ".join(errors))
    first_errors = errors[:4]
    last_errors = errors[-4:]
    merged = first_errors + [f"... skipped {len(errors) - 8} errors ..."] + last_errors
    raise RuntimeError(" ; ".join(merged))


def apply_rerank_results(relations: pd.DataFrame, pair_results: dict[str, dict[str, Any]]) -> tuple[pd.DataFrame, dict[str, int], dict[str, int]]:
    updated = relations.copy()
    targeted = target_mask(updated)
    updated["pair_key"] = updated.apply(lambda row: build_pair_key(row["source_person_id"], row["target_person_id"]), axis=1)

    before_counts = updated["raw_relation_type"].astype(str).value_counts().to_dict()
    auto_formal = 0
    review_count = 0
    hidden_count = 0

    for idx, row in updated.iterrows():
        raw_relation_type = text(row["raw_relation_type"]) or text(row["standard_relation_type"])
        current_relation_type = text(row.get("final_relation_type", "")) or text(row.get("standard_relation_type", "")) or raw_relation_type
        updated.at[idx, "raw_relation_type"] = raw_relation_type
        if row["pair_key"] not in pair_results:
            updated.at[idx, "final_relation_type"] = current_relation_type
            updated.at[idx, "display_status"] = "review" if bool(targeted.loc[idx]) else "formal"
            if bool(targeted.loc[idx]):
                updated.at[idx, "llm_reason"] = "LLM 批处理中未返回该人物对，保留待审核。"
            continue

        result = pair_results[row["pair_key"]]
        suggested = text(result.get("llm_suggested_relation_type"))
        reason = text(result.get("llm_reason"))
        confidence = float(result.get("llm_confidence", 0.0) or 0.0)
        evidence_strength = text(result.get("evidence_strength")).lower()

        display_status = "review"
        final_relation_type = current_relation_type
        if confidence >= 0.85 and evidence_strength == "strong" and suggested not in HIDDEN_LABELS:
            final_relation_type = suggested
            display_status = "formal"
            auto_formal += 1
        elif confidence < 0.60 or suggested in HIDDEN_LABELS:
            final_relation_type = suggested or raw_relation_type
            display_status = "hidden"
            hidden_count += 1
        else:
            final_relation_type = suggested or raw_relation_type
            display_status = "review"
            review_count += 1

        updated.at[idx, "llm_suggested_relation_type"] = suggested
        updated.at[idx, "final_relation_type"] = final_relation_type
        updated.at[idx, "llm_reason"] = reason
        updated.at[idx, "llm_confidence"] = round(confidence, 4)
        updated.at[idx, "display_status"] = display_status

    updated.drop(columns=["pair_key"], inplace=True)
    after_counts = updated["final_relation_type"].astype(str).value_counts().to_dict()
    status_counts = {"formal": auto_formal, "review": review_count, "hidden": hidden_count}
    return updated, before_counts, after_counts | status_counts


def validate_app_runtime() -> dict[str, str]:
    command = [
        sys.executable,
        "-m",
        "streamlit",
        "run",
        "app.py",
        "--server.headless",
        "true",
        "--server.address",
        "127.0.0.1",
        "--server.port",
        "8522",
    ]
    process = subprocess.Popen(
        command,
        cwd=APP_DIR,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
        bufsize=1,
    )
    lines: list[str] = []
    success = False
    try:
        deadline = time.time() + 30
        while time.time() < deadline:
            line = process.stdout.readline() if process.stdout else ""
            if line:
                lines.append(line.rstrip())
                if "You can now view your Streamlit app in your browser" in line or "Local URL:" in line or "URL:" in line:
                    success = True
                    break
            elif process.poll() is not None:
                break
            else:
                time.sleep(0.2)
    finally:
        if process.poll() is None:
            process.terminate()
            try:
                process.wait(timeout=8)
            except subprocess.TimeoutExpired:
                process.kill()
        if process.stdout:
            rest = process.stdout.read()
            if rest:
                lines.extend(rest.splitlines())
    return {"success": "yes" if success else "no", "details": "\n".join(lines[-30:])}


def write_report(
    *,
    updated_relations: pd.DataFrame,
    before_counts: dict[str, int],
    app_validation: dict[str, str],
) -> None:
    report_relations = updated_relations[rerank_scope_mask(updated_relations)].copy()
    report_relations["pair_key"] = report_relations.apply(lambda row: build_pair_key(row["source_person_id"], row["target_person_id"]), axis=1)
    targeted_pair_count = int(report_relations["pair_key"].nunique())
    targeted_row_count = int(len(report_relations))
    status_counts = report_relations["display_status"].astype(str).value_counts().to_dict()
    changes = (
        report_relations[report_relations["raw_relation_type"].astype(str) != report_relations["final_relation_type"].astype(str)]
        .groupby(["raw_relation_type", "final_relation_type"])
        .size()
        .reset_index(name="count")
        .sort_values("count", ascending=False)
    )
    lines = [
        "# llm_relation_rerank_report",
        "",
        f"- 重判总数（人物对）: {targeted_pair_count}",
        f"- 重判总数（关系记录）: {targeted_row_count}",
        f"- 自动转正数量: {status_counts.get('formal', 0)}",
        f"- review 数量: {status_counts.get('review', 0)}",
        f"- hidden 数量: {status_counts.get('hidden', 0)}",
        "",
        "## 各关系类型变化",
    ]
    if changes.empty:
        lines.append("- 无类型变化。")
    else:
        for _, row in changes.iterrows():
            lines.append(f"- {row['raw_relation_type']} -> {row['final_relation_type']}: {int(row['count'])}")

    lines.extend(
        [
            "",
            "## 最终展示类型分布",
        ]
    )
    final_counts = report_relations["final_relation_type"].astype(str).value_counts().to_dict()
    for label, count in final_counts.items():
        lines.append(f"- {label}: {count}")

    lines.extend(
        [
            "",
            "## 原始类型分布",
        ]
    )
    for label, count in before_counts.items():
        lines.append(f"- {label}: {count}")

    lines.extend(
        [
            "",
            "## app.py 运行结果",
            f"- 成功运行: {app_validation['success']}",
            "```text",
            app_validation["details"] or "(no output)",
            "```",
        ]
    )
    REPORT_PATH.write_text("\n".join(lines), encoding="utf-8")


def main() -> None:
    openai_base_url = text(os.getenv("OPENAI_BASE_URL") or os.getenv("RELATION_RERANK_BASE_URL"))
    openai_api_key = text(os.getenv("OPENAI_API_KEY") or os.getenv("RELATION_RERANK_API_KEY"))
    anthropic_base_url = text(os.getenv("ANTHROPIC_BASE_URL"))
    anthropic_auth_token = text(os.getenv("ANTHROPIC_AUTH_TOKEN"))
    batch_size = int(text(os.getenv("RELATION_RERANK_BATCH_SIZE")) or DEFAULT_BATCH_SIZE)
    sleep_seconds = float(text(os.getenv("RELATION_RERANK_SLEEP_SECONDS")) or DEFAULT_SLEEP_SECONDS)
    max_pairs = int(text(os.getenv("RELATION_RERANK_MAX_PAIRS")) or "0")
    report_only = text(os.getenv("RELATION_RERANK_REPORT_ONLY")).lower() in {"1", "true", "yes", "on"}

    persons, relations, sources = load_frames()
    relations = ensure_columns(relations)
    scope_mask = rerank_scope_mask(relations)
    if report_only:
        app_validation = validate_app_runtime()
        write_report(
            updated_relations=relations,
            before_counts=relations[scope_mask]["raw_relation_type"].astype(str).value_counts().to_dict(),
            app_validation=app_validation,
        )
        relations.to_csv(RELATIONS_PATH, index=False, encoding="utf-8-sig")
        print("REPORT_ONLY")
        print(f"REPORT={REPORT_PATH}")
        return
    pair_payloads, _ = aggregate_pairs(relations, persons, sources)
    if max_pairs > 0:
        pair_payloads = pair_payloads[:max_pairs]
    targeted_pair_count = len(pair_payloads)
    targeted_row_count = int(target_mask(relations).sum())
    if targeted_pair_count == 0:
        app_validation = validate_app_runtime()
        write_report(
            updated_relations=relations,
            before_counts=relations[scope_mask]["raw_relation_type"].astype(str).value_counts().to_dict(),
            app_validation=app_validation,
        )
        relations.to_csv(RELATIONS_PATH, index=False, encoding="utf-8-sig")
        print("NO_TARGET_RELATIONS")
        return

    pair_results: dict[str, dict[str, Any]] = {}
    total_batches = math.ceil(targeted_pair_count / batch_size)
    for index in range(total_batches):
        batch = pair_payloads[index * batch_size : (index + 1) * batch_size]
        if openai_base_url and openai_api_key:
            response = call_openai_batch(batch, openai_base_url, openai_api_key, get_openai_models())
        else:
            response = call_anthropic_batch(batch, anthropic_base_url, anthropic_auth_token, get_models())
        for item in response.get("results", []):
            pair_key = text(item.get("pair_key"))
            if not pair_key:
                continue
            pair_results[pair_key] = {
                "llm_suggested_relation_type": text(item.get("llm_suggested_relation_type")),
                "llm_confidence": float(item.get("llm_confidence", 0.0) or 0.0),
                "evidence_strength": text(item.get("evidence_strength")).lower(),
                "llm_reason": text(item.get("llm_reason"))[:600],
            }
        print(f"BATCH {index + 1}/{total_batches} DONE model={response.get('_model','')} endpoint={response.get('_endpoint','')}")
        time.sleep(sleep_seconds)

    updated_relations, before_counts, _ = apply_rerank_results(relations, pair_results)
    updated_relations.to_csv(RELATIONS_PATH, index=False, encoding="utf-8-sig")
    app_validation = validate_app_runtime()
    write_report(
        updated_relations=updated_relations,
        before_counts=updated_relations[rerank_scope_mask(updated_relations)]["raw_relation_type"].astype(str).value_counts().to_dict(),
        app_validation=app_validation,
    )
    print(f"TARGETED_PAIRS={targeted_pair_count}")
    print(f"TARGETED_ROWS={targeted_row_count}")
    print(f"UPDATED_FILE={RELATIONS_PATH}")
    print(f"REPORT={REPORT_PATH}")


if __name__ == "__main__":
    main()
