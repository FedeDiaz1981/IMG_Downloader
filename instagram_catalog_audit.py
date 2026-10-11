"""Audit Pinto Fruta catalog presence in Instagram posts using the signed-in Edge session."""
from __future__ import annotations

import base64
import csv
import json
import re
import sys
from datetime import datetime
from pathlib import Path

import requests
from playwright.sync_api import sync_playwright

import app

PROFILE_URL = "https://www.instagram.com/distribuidora_pintofruta/"


def analyze_post(image_path: Path, text: str, api_key: str) -> list[str]:
    prompt = (
        "Identifica productos comercializados visibles en esta publicacion de Instagram de una distribuidora. "
        "Usa la imagen y el texto. Devuelve JSON estricto: {\"products\":[\"marca producto presentacion\"]}. "
        "No inventes productos: si no hay uno reconocible devuelve una lista vacia.\n\n"
        f"Texto publicado: {text[:1800]}"
    )
    data = image_path.read_bytes()
    payload = {"contents": [{"parts": [{"text": prompt}, {"inline_data": {
        "mime_type": "image/jpeg", "data": base64.b64encode(data).decode("ascii")
    }}]}], "generationConfig": {"temperature": 0, "responseMimeType": "application/json"}}
    for model in ("gemini-3.5-flash-lite", "gemini-2.5-flash-lite"):
        response = requests.post(
            f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent",
            params={"key": api_key}, json=payload, timeout=45,
        )
        if response.status_code in (400, 404):
            continue
        response.raise_for_status()
        answer = response.json()["candidates"][0]["content"]["parts"][0]["text"]
        return [str(value) for value in json.loads(answer).get("products", []) if str(value).strip()]
    return []


def catalog_match(product: str, catalog: list[dict[str, object]]) -> dict[str, object] | None:
    wanted = set(app.product_search_tokens(product))
    best: tuple[float, dict[str, object]] | None = None
    for item in catalog:
        available = set(app.product_search_tokens(f"{item['brand']} {item['name']}"))
        if not wanted or not available:
            continue
        score = len(wanted & available) / len(wanted)
        if best is None or score > best[0]:
            best = score, item
    return best[1] if best and best[0] >= 0.75 else None


def main() -> None:
    api_key = app.get_config_value("GEMINI_API_KEY")
    if not api_key:
        raise RuntimeError("Falta GEMINI_API_KEY")
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    work_dir = app.JOB_DIR / f"instagram-audit-{stamp}"
    work_dir.mkdir()
    log_path = work_dir / "progreso.txt"
    def log(message: str) -> None:
        print(message, flush=True)
        with log_path.open("a", encoding="utf-8") as file:
            file.write(message + "\n")

    with sync_playwright() as playwright:
        browser = playwright.chromium.connect_over_cdp(app.PINFOFRUTA_DEBUG_URL)
        context = browser.contexts[0]
        profile = context.new_page()
        profile.goto(PROFILE_URL, wait_until="domcontentloaded", timeout=30_000)
        profile.wait_for_timeout(1_500)
        urls: set[str] = set()
        stable = 0
        for _ in range(100):
            before = len(urls)
            urls.update(profile.locator("a[href*='/p/'], a[href*='/reel/']").evaluate_all("els => els.map(e => e.href)"))
            profile.mouse.wheel(0, 1800)
            profile.wait_for_timeout(350)
            stable = stable + 1 if len(urls) == before else 0
            if stable >= 8:
                break
        admin = context.new_page()
        admin.set_default_timeout(15_000)
        catalog = app.pintofruta_admin_products(admin)
        log(f"Publicaciones: {len(urls)}. Catalogo: {len(catalog)}.")
        seen_ids: set[int] = set()
        detected: list[tuple[str, str, str]] = []
        post_page = context.new_page()
        for index, url in enumerate(sorted(urls), start=1):
            try:
                post_page.goto(url, wait_until="domcontentloaded", timeout=30_000)
                post_page.wait_for_timeout(700)
                text = post_page.locator("body").inner_text()[:4000]
                shot = work_dir / f"{index:03}.jpg"
                article = post_page.locator("article")
                (article if article.count() else post_page.locator("body")).screenshot(path=str(shot), type="jpeg", quality=75)
                for product in analyze_post(shot, text, api_key):
                    match = catalog_match(product, catalog)
                    if match:
                        seen_ids.add(int(match["id"]))
                        detected.append((url, product, str(match["name"])))
                log(f"{index}/{len(urls)} procesada")
            except Exception as exc:
                log(f"{index}/{len(urls)} error: {str(exc)[:120]}")
        browser.close()

    report = app.JOB_DIR / f"catalogo-no-detectado-instagram-{stamp}.csv"
    with report.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(["ID Pinto Fruta", "Marca", "Producto", "Resultado"])
        for item in catalog:
            if int(item["id"]) not in seen_ids:
                writer.writerow([item["id"], item["brand"], item["name"], "No detectado en publicaciones analizadas"])
    detected_report = work_dir / "productos_detectados.csv"
    with detected_report.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file); writer.writerow(["Publicacion", "Detectado por IA", "Producto catalogo"]); writer.writerows(detected)
    log(f"FINAL: detectados={len(seen_ids)}; reporte={report}")


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        print(f"ERROR: {error}", file=sys.stderr, flush=True)
        raise
