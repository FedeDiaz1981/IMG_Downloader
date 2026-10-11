"""Deactivate, without deleting, products absent from the Instagram audit."""
from __future__ import annotations

import csv
import os
import re
from datetime import datetime
from pathlib import Path

from playwright.sync_api import sync_playwright

import app


def product_page(context):
    """Return the fullest live Pinto Fruta product-list tab."""
    candidates = []
    for item in context.pages:
        if "pintofruta.com" not in item.url or item.is_closed():
            continue
        try:
            if "LISTAS / PRODUCTOS" in item.locator("body").inner_text(timeout=5_000):
                candidates.append(item)
        except Exception:
            continue
    def catalog_total(item):
        match = re.search(r"Productos\s+(\d+)", item.locator("body").inner_text())
        return int(match.group(1)) if match else 999_999

    page = min(candidates, key=catalog_total, default=None)
    if page is None:
        page = next(
            (item for item in context.pages if "pintofruta.com" in item.url and not item.is_closed()),
            None,
        )
    if page is None:
        raise RuntimeError("No se encontro una pestana de Pinto Fruta abierta")
    page.set_default_timeout(15_000)
    if "LISTAS / PRODUCTOS" not in page.locator("body").inner_text():
        page.get_by_role("button", name="Productos").first.click()
        page.wait_for_function("() => document.body.innerText.includes('LISTAS / PRODUCTOS')")
    return page


def main() -> None:
    resume_source = os.environ.get("IMG_DOWNLOADER_RESUME_SOURCE")
    source = (
        Path(resume_source)
        if resume_source
        else max(app.JOB_DIR.glob("catalogo-no-detectado-instagram-*.csv"), key=lambda path: path.stat().st_mtime)
    )
    with source.open(encoding="utf-8-sig", newline="") as file:
        targets = list(csv.DictReader(file))
    if resume_source:
        targets = [item for item in targets if item.get("Resultado", "").startswith("Error:")]
    limit = int(os.environ.get("IMG_DOWNLOADER_LIMIT", "0"))
    if limit:
        targets = targets[:limit]

    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    audit = app.JOB_DIR / f"productos-inactivados-instagram-{stamp}.csv"
    results: list[list[str]] = []

    with sync_playwright() as playwright:
        browser = playwright.chromium.connect_over_cdp(app.PINFOFRUTA_DEBUG_URL, timeout=15_000)
        context = browser.contexts[0]
        initial_page = product_page(context)
        cancel = initial_page.get_by_role("button", name="Cancelar")
        if cancel.count() and cancel.last.is_visible():
            print("Cerrando una ventana de edicion pendiente antes de iniciar", flush=True)
            cancel.last.click()
            initial_page.wait_for_timeout(500)
        for index, item in enumerate(targets, start=1):
            product_id = item["ID Pinto Fruta"]
            sku = item.get("SKU", "")
            try:
                page = product_page(context)
                search = page.locator("input[placeholder*='SKU']").first
                search.fill("")
                page.wait_for_timeout(800)
                search.fill(sku or item["Producto"])
                page.wait_for_timeout(1_000)
                target_row = None
                for row_index in range(page.locator("tbody tr").count()):
                    row = page.locator("tbody tr").nth(row_index)
                    row_id = re.sub(r"\D", "", row.locator("td").nth(1).inner_text())
                    if row_id == product_id:
                        target_row = row
                        break
                if target_row is None:
                    results.append(["No encontrado en catalogo activo", product_id, sku, item["Marca"], item["Producto"]])
                    print(f"{index}/{len(targets)} fuera del catalogo activo {product_id}", flush=True)
                    continue
                target_row.get_by_role("button", name="Editar").click()
                active = page.locator("label").filter(
                    has_text=re.compile(r"^ACTIVO$", re.I)
                ).first.locator("input[type='checkbox']")
                active.wait_for(state="visible")
                if not active.is_checked():
                    page.get_by_role("button", name="Cancelar").last.click()
                    results.append(["Ya estaba inactivo", product_id, sku, item["Marca"], item["Producto"]])
                    print(f"{index}/{len(targets)} ya inactivo {product_id}", flush=True)
                    continue
                active.uncheck()
                page.get_by_role("button", name="Guardar").last.click()
                active.wait_for(state="hidden", timeout=20_000)
                results.append(["Inactivado", product_id, sku, item["Marca"], item["Producto"]])
                print(f"{index}/{len(targets)} inactivado {product_id}", flush=True)
            except Exception as exc:
                results.append([f"Error: {str(exc)[:180]}", product_id, sku, item["Marca"], item["Producto"]])
                print(f"{index}/{len(targets)} ERROR {product_id}: {exc}", flush=True)

    with audit.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(["Resultado", "ID Pinto Fruta", "SKU", "Marca", "Producto"])
        writer.writerows(results)
    print(f"FINAL: {audit}", flush=True)


if __name__ == "__main__":
    main()
