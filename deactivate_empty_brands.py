"""Deactivate active Pinto Fruta brands that have no active products."""
from __future__ import annotations

import csv
import os
import re
from datetime import datetime

from playwright.sync_api import sync_playwright

import app


def open_section(context, section: str):
    page = context.new_page()
    page.set_default_timeout(15_000)
    page.goto(app.PINFOFRUTA_ADMIN_URL, wait_until="commit", timeout=15_000)
    page.wait_for_timeout(2_000)
    page.get_by_role("button", name=section).first.click()
    page.wait_for_timeout(1_000)
    return page


def table_pages(page) -> list[list[str]]:
    records: list[list[str]] = []
    seen_pages: set[str] = set()
    while True:
        rows = page.locator("tbody tr").evaluate_all(
            "els => els.map(row => [...row.querySelectorAll('td')].map(cell => cell.innerText.trim()))"
        )
        signature = "|".join(rows[0]) if rows else "empty"
        if signature in seen_pages:
            break
        seen_pages.add(signature)
        records.extend(rows)
        next_button = page.get_by_role("button", name="Siguiente")
        if next_button.is_disabled():
            break
        next_button.click()
        page.wait_for_timeout(700)
    return records


def brand_row_by_id(page, brand_id: str, brand_name: str):
    search = page.locator("input[placeholder*='SKU']").first
    search.fill(brand_id or brand_name)
    page.wait_for_timeout(700)
    for index in range(page.locator("tbody tr").count()):
        row = page.locator("tbody tr").nth(index)
        row_id = row.locator("td").nth(1).inner_text().strip()
        if row_id == brand_id:
            return row
    return None


def main() -> None:
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    report_path = app.JOB_DIR / f"marcas-sin-productos-activos-{stamp}.csv"
    dry_run = os.environ.get("IMG_DOWNLOADER_DRY_RUN") == "1"
    limit = int(os.environ.get("IMG_DOWNLOADER_LIMIT", "0"))
    results: list[list[str]] = []

    with sync_playwright() as playwright:
        browser = playwright.chromium.connect_over_cdp(app.PINFOFRUTA_DEBUG_URL, timeout=15_000)
        context = browser.contexts[0]
        product_page = open_section(context, "Productos")
        products = table_pages(product_page)
        active_product_brands = {
            app.normalize_text(row[4]) for row in products if len(row) > 4 and row[4].strip()
        }

        brand_page = open_section(context, "Marcas")
        brands = table_pages(brand_page)
        candidates = []
        for row in brands:
            if len(row) < 6:
                continue
            brand_id = row[1].strip()
            brand_name = row[3].strip()
            active_status = app.normalize_text(row[5])
            if not brand_id or not brand_name or active_status != "activo":
                continue
            if app.normalize_text(brand_name) not in active_product_brands:
                candidates.append((brand_id, brand_name))
        if limit:
            candidates = candidates[:limit]

        print(f"Productos activos: {len(products)}", flush=True)
        print(f"Marcas activas sin productos activos: {len(candidates)}", flush=True)
        for brand_id, brand_name in candidates:
            if dry_run:
                results.append(["Pendiente de inactivar", brand_id, brand_name])
                print(f"Pendiente: {brand_name}", flush=True)
                continue
            row = brand_row_by_id(brand_page, brand_id, brand_name)
            if row is None:
                results.append(["Error: marca no encontrada", brand_id, brand_name])
                continue
            try:
                row.get_by_role("button", name="Editar").click()
                active = brand_page.locator("label").filter(
                    has_text=re.compile(r"^ACTIVA$", re.I)
                ).first.locator("input[type='checkbox']")
                active.wait_for(state="visible")
                if not active.is_checked():
                    brand_page.get_by_role("button", name="Cancelar").last.click()
                    results.append(["Ya estaba inactiva", brand_id, brand_name])
                    continue
                active.uncheck()
                brand_page.get_by_role("button", name="Guardar").last.click()
                active.wait_for(state="hidden", timeout=20_000)
                results.append(["Inactivada", brand_id, brand_name])
                print(f"Inactivada: {brand_name}", flush=True)
            except Exception as exc:
                results.append([f"Error: {str(exc)[:180]}", brand_id, brand_name])
                print(f"Error {brand_name}: {exc}", flush=True)

    with report_path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(["Resultado", "ID", "Marca"])
        writer.writerows(results)
    print(f"FINAL: {report_path}", flush=True)


if __name__ == "__main__":
    main()
