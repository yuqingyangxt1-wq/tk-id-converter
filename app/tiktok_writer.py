"""Write product rows into the TikTok Shop Indonesia batch upload template.

The template (assets/batch-product-source.xlsx) is the OFFICIAL empty template
provided by TikTok Shop Indonesia (40 cols, Indonesian values throughout).
We only write data into the 'Template' sheet starting at row 2, leaving the
rest of the workbook (Instruction, Image, Example, HiddenStyle, HiddenAttr,
Category, Brand, ShippingInsurance, Condition, etc.) untouched.

ROW STRATEGY (v1.0.10 — OFFICIAL Indonesia template):
    The EasyBoss source table has one row per (color × size) SKU. We expand
    each SKU into its own TikTok listing row (each row = one variant of one
    product).

    v1.0.10 — 使用官方 TikTok Shop ID 模板 (替换 EasyBoss 参考):
    1. Template 40 列 (无 shipping_insurance 列 — 官方模板没有)
    2. HiddenAttr 所有合法值 = 印尼语 (Katun, Polos, Musim semi, Atletis, ...)
       之前 v1.0.9 抄 EasyBoss 参考是英文值 (Cotton, Plain, All seasons) — 错！
       EasyBoss 工具用的是英译版模板，Tiktok 后台要求印尼语值
    3. Category 用印尼语 ("Atasan Pria/T-shirt" 而不是 "Men's Tops/T-shirts")
    4. Brand = "Tidak ada merek" (印尼语 "没有品牌" 而不是 "No brand")
    5. ShippingInsurance = "Opsional" 而不是 "Optional"
    6. 文件不输出 shipping_insurance 列 (官方模板没有这一列)
"""
from __future__ import annotations

import shutil
from pathlib import Path
from typing import Any

import openpyxl

from .config import get_assets_dir
from .source_reader import Product, Variant


TEMPLATE_FILENAME = "batch-product-source.xlsx"

# v1.0.10: 40 列官方模板 (无 shipping_insurance)
TIKTOK_COLUMNS: list[str] = [
    "category",
    "brand",
    "product_name",
    "product_description",
    "main_image",
    "image_2", "image_3", "image_4", "image_5",
    "image_6", "image_7", "image_8", "image_9",
    "property_name_1",
    "property_value_1",
    "property_1_image",
    "property_name_2",
    "property_value_2",
    "parcel_weight",
    "parcel_length",
    "parcel_width",
    "parcel_height",
    "delivery",
    "price",
    "quantity",
    "seller_sku",
    "minimum_order_quantity",
    "size_chart",
    "cod",
    "product_property/100157",
    "product_property/100198",
    "product_property/100393",
    "product_property/100395",
    "product_property/100397",
    "product_property/100398",
    "product_property/100399",
    "product_property/100400",
    "product_property/100401",
    "product_property/100403",
]


# v1.0.10: 官方 Category 用印尼语路径
_ID_CATEGORY_TRANSLATION: dict[str, str] = {
    # 男装上衣
    "Pakaian & Pakaian Dalam Pria>Atasan Pria>T-shirt": "Atasan Pria/T-shirt",
    "Pakaian & Pakaian Dalam Pria>Atasan Pria>Kaus Polo": "Atasan Pria/Kaus Polo",
    "Pakaian & Pakaian Dalam Pria>Atasan Pria>Kemeja": "Atasan Pria/Kemeja",
    "Pakaian & Pakaian Dalam Pria>Atasan Pria>Rajut": "Atasan Pria/Pakaian Rajut",
    "Pakaian & Pakaian Dalam Pria>Atasan Pria>Hoodie & Sweatshirt": "Atasan Pria/Hoodie & Jumper",
    "Pakaian & Pakaian Dalam Pria>Atasan Pria>Rompi & Gilet": "Atasan Pria/Rompi Waistcoat & Gilet",
    "Pakaian & Pakaian Dalam Pria>Atasan Pria>Jaket & Mantel": "Atasan Pria/Jaket & Mantel",
    # 男装下装
    "Pakaian & Pakaian Dalam Pria>Celana Pria>Celana Pendek": "Bawahan Pria/Celana pendek",
    "Pakaian & Pakaian Dalam Pria>Celana Pria>Jeans": "Bawahan Pria/Jeans",
    "Pakaian & Pakaian Dalam Pria>Celana Pria>Celana Panjang": "Bawahan Pria/Celana Pria",
    # 男士套装
    "Setelan Pria>Set Pakaian Pria": "Setelan & Overall Pria/Set Pakaian Pria",
    "Setelan Pria>Setelan Pria": "Setelan & Overall Pria/Setelan Resmi",
    "Setelan Pria>Overall": "Setelan & Overall Pria/Overall",
    # 男士内衣袜子
    "Pakaian Dalam & Kaus Kaki Pria>Pakaian Dalam Pria": "Pakaian Dalam Pria/Pakaian Dalam",
    "Pakaian Dalam & Kaus Kaki Pria>Tank Top & Pakaian Dalam": "Pakaian Dalam Pria/Pakaian Dalam",
    "Pakaian Dalam & Kaus Kaki Pria>Pakaian Dalam Hangat": "Pakaian Dalam Pria/Pakaian Dalam Termal",
    "Pakaian Dalam & Kaus Kaki Pria>Kaus Kaki": "Pakaian Dalam Pria/Kaus kaki",
    # 男士睡衣
    "Pakaian Tidur & Pakaian Santai Pria>Piyama & Loungewear": "Baju Tidur dan Baju Santai Pria/Piyama",
    "Pakaian Tidur & Pakaian Santai Pria>Robe Pria": "Baju Tidur dan Baju Santai Pria/Kimono Mandi & Rias",
    "Pakaian Tidur & Pakaian Santai Pria>Nightshirt": "Baju Tidur dan Baju Santai Pria/Piyama Midi",
    "Pakaian Tidur & Pakaian Santai Pria>Piyama Pria": "Baju Tidur dan Baju Santai Pria/Piama Terusan Pria",
    # 男士特殊场合
    "Pakaian Acara Khusus Pria>Kostum": "Pakaian Khusus Pria/Kostum & Aksesoris",
    "Pakaian Acara Khusus Pria>Pakaian Kerja": "Pakaian Khusus Pria/Pakaian Kerja & Seragam",
    "Pakaian Acara Khusus Pria>Pakaian Tradisional": "Pakaian Khusus Pria/Baju Tradisional",
}


def translate_category(raw: str) -> str:
    """Translate an EasyBoss ID Indonesian category path into OFFICIAL TikTok
    Indonesia template Indonesian category name."""
    if not raw:
        return raw
    return _ID_CATEGORY_TRANSLATION.get(raw, raw)


# v1.0.10: 变体名保持英文（property_value 用印尼语 color 来源用印尼语）
VAR_NAME_TRANSLATIONS: dict[str, str] = {
    "颜色": "Color",
    "顏色": "Color",
    "尺码": "Size",
    "尺碼": "Size",
    "尺寸": "Size",
    "规格": "Specification",
    "Color": "Color",
    "color": "Color",
    "colour": "Color",
    "Size": "Size",
    "size": "Size",
}


def _translate_var_name(name: str) -> str:
    if not name:
        return ""
    key = _clean_str(name)
    if key in VAR_NAME_TRANSLATIONS:
        return VAR_NAME_TRANSLATIONS[key]
    return key


def template_path() -> Path:
    return get_assets_dir() / TEMPLATE_FILENAME


def _kg_to_grams(v: Any) -> Any:
    if v is None or v == "":
        return None
    try:
        f = float(v)
        return int(round(f * 1000))
    except (TypeError, ValueError):
        return v


def _clean_str(v: Any) -> str:
    if v is None:
        return ""
    return str(v).strip()


def _to_number(v: Any) -> Any:
    if v is None or v == "":
        return None
    if isinstance(v, bool):
        return int(v)
    if isinstance(v, (int, float)):
        return v
    s = str(v).strip()
    if not s:
        return None
    for ch in [",", " ", "₱", "$", "￥", "¥", "PHP", "php", "kg", "KG", "g", "G", "cm", "CM"]:
        s = s.replace(ch, "")
    try:
        f = float(s)
        if f.is_integer():
            return int(f)
        return f
    except ValueError:
        return v


OutputRow = dict[str, Any]


def _resolve_common_fields(
    product: Product,
    settings: dict[str, Any],
) -> dict[str, Any]:
    """Resolve fields that are shared across all rows of a product."""
    # v1.0.10: 默认 brand = "Tidak ada merek"（印尼语 "没有品牌"）
    brand = (
        settings["brand_value"]
        if settings.get("brand_enabled") and settings.get("brand_value")
        else (product.brand or "")
    )
    if not brand:
        brand = "Tidak ada merek"

    price = (
        settings["price_value"]
        if settings.get("price_enabled")
        else None
    )
    quantity = (
        settings["quantity_value"]
        if settings.get("quantity_enabled")
        else None
    )
    cod = settings["cod_value"] if settings.get("cod_enabled") else ""
    category = (
        settings["category_value"]
        if settings.get("category_enabled")
        else product.category
    )
    category = translate_category(category)
    description = (
        settings["description_value"]
        if settings.get("description_enabled")
        else product.description
    )
    size_chart = (
        settings["size_chart_value"]
        if settings.get("size_chart_enabled")
        else product.size_chart
    )

    if settings.get("parcel_enabled"):
        weight = settings.get("parcel_weight_value")
        length = settings.get("parcel_length_value")
        width = settings.get("parcel_width_value")
        height = settings.get("parcel_height_value")
    else:
        weight = _kg_to_grams(product.parcel_weight_kg)
        length = product.parcel_length
        width = product.parcel_width
        height = product.parcel_height

    return {
        "brand": _clean_str(brand),
        "category": _clean_str(category),
        "description": _clean_str(description),
        "size_chart": _clean_str(size_chart),
        "weight": weight,
        "length": length,
        "width": width,
        "height": height,
        "cod": cod,
        "quantity": quantity,
        "price_override": price,
        "delivery": _clean_str(settings.get("delivery_value", "")),
    }


def _build_row_for_variant(
    product: Product,
    variant: Variant,
    settings: dict[str, Any],
    common: dict[str, Any],
    copy_idx: int = 0,
    copy_suffix: str = "",
    apply_suffix_to_first: bool = False,
) -> OutputRow:
    use_suffix = (copy_idx > 0) or apply_suffix_to_first

    title_prefix = settings["title_prefix"] if settings.get("title_prefix_enabled") else ""
    suffix = copy_suffix if use_suffix else ""
    title = f"{title_prefix}{product.product_name}{suffix}".strip()
    title = (
        title.replace("\u2013", " - ")
             .replace("\u2014", " - ")
             .replace("\u2019", "'")
             .replace("\u201c", '"')
             .replace("\u201d", '"')
    )
    if len(title) > 100:
        title = title[:97].rstrip() + "..."

    var1_name = _translate_var_name(product.var1_name or "颜色")
    var1_value = _clean_str(variant.var1)  # 印尼语颜色原值透传
    var2_name = _translate_var_name(product.var2_name or "尺码")

    if variant.var2:
        var2_value = _clean_str(variant.var2)
    elif settings.get("fill_sizes_enabled"):
        std = settings.get("standard_sizes") or "S,M,L,XL,2XL,3XL"
        var2_value = std
    else:
        var2_value = ""

    var1_image = _clean_str(variant.sku_image)
    images = list(product.images[:9])
    while len(images) < 9:
        images.append("")
    images = [_clean_str(i) for i in images]

    price_value = _to_number(variant.price) if variant.price not in (None, "") else _to_number(common["price_override"])

    quantity_value = (
        _to_number(variant.stock) if variant.stock not in (None, "") else _to_number(common["quantity"])
    )

    import hashlib as _hl
    base_sku = _clean_str(variant.platform_sku)
    if not base_sku:
        prod_hash = _hl.md5(product.product_name.encode("utf-8")).hexdigest()[:6].upper()
        size_part = (f"-{var1_value}" if var1_value else "") + (f"-{var2_value}" if var2_value else "")
        base_sku = f"ID{prod_hash}{size_part}"
    seller_sku = base_sku + copy_suffix if copy_suffix and use_suffix else base_sku

    row: OutputRow = {col: "" for col in TIKTOK_COLUMNS}
    row["category"] = common["category"]
    row["brand"] = common["brand"]
    row["product_name"] = title
    row["product_description"] = common["description"]
    row["main_image"] = var1_image if var1_image else images[0] if images else ""
    for i, img in enumerate(images[:9]):
        col = "main_image" if i == 0 else f"image_{i + 1}"
        if i == 0 and var1_image:
            continue
        row[col] = img
    row["property_name_1"] = var1_name
    row["property_value_1"] = var1_value
    row["property_1_image"] = var1_image
    row["property_name_2"] = var2_name
    row["property_value_2"] = var2_value
    row["parcel_weight"] = _to_number(common["weight"])
    row["parcel_length"] = _to_number(common["length"])
    row["parcel_width"] = _to_number(common["width"])
    row["parcel_height"] = _to_number(common["height"])
    row["delivery"] = common["delivery"]
    row["price"] = price_value
    row["quantity"] = quantity_value
    row["seller_sku"] = seller_sku
    row["size_chart"] = common["size_chart"]
    row["cod"] = common["cod"]
    row["minimum_order_quantity"] = settings.get("minimum_order_quantity_value", 1) or 1
    # v1.0.10: 不再写 shipping_insurance (官方模板没有这一列)
    category = common.get("category", "")
    for prop_id, default_val in _get_property_fallbacks(category).items():
        row[prop_id] = default_val
    return row


# v1.0.10: PREFERRED 默认值用印尼语（参考官方 HiddenAttr 表）
_PROPERTY_FALLBACK_CACHE: dict[str, dict[str, str]] | None = None


# v1.0.11: 之前硬编码 PREFERRED 默认值错的——T-shirt 在不同 prop_id 的合法值
# 是 "V-Neck" "Lengan pendek" "Musim semi" "Atletis" "Slim-fit" 等，
# 不是我之前以为的 "Polos, Semua musim, Dasar, Pas"。
# PREFERRED 完全删除，让 HiddenAttr 模板告诉你该填什么。
# HiddenAttr 9 对列 (C1-C18) 对应 Template C31-C39 (prop_id 100157~100401)，
# Template C40 (100403) 没有 HiddenAttr pair → 留空。
# 列映射:
#   pair 0 (C1/C2)  → Template C31 (100157 Material)
#   pair 1 (C3/C4)  → Template C32 (100198 Pattern)
#   pair 2 (C5/C6)  → Template C33 (100393 Neckline)
#   pair 3 (C7/C8)  → Template C34 (100395 Sleeve length)
#   pair 4 (C9/C10) → Template C35 (100397 Season)
#   pair 5 (C11/C12)→ Template C36 (100398 Style)
#   pair 6 (C13/C14)→ Template C37 (100399 Fit type)
#   pair 7 (C15/C16)→ Template C38 (100400 Stretch)
#   pair 8 (C17/C18)→ Template C39 (100401 Care instructions)
# v1.0.14: PREFERRED 重新启用，但只填 HiddenAttr 合法下拉列表里实际存在的值。
# 之前 v1.0.13 我误以为 "Cowl Neck" / "Semua musim" 不在 HiddenAttr 合法列表里
# （只看了 R2 header 行，没展开 R225-R251 / R41-R45 的同 pair 多类目行）→ 删除
# PREFERRED 让 HiddenAttr 取首选项 V-Neck / Musim semi，导致用户实际 Cowl Neck
# 款 T-shirt 的 Neckline 列填错（虽然合法但不匹配产品）。修正：HiddenAttr 已
# 确认这两个值都合法（R243 Cowl Neck, R45 Semua musim），PREFERRED 优先。
_PROPERTY_PREFERRED_DEFAULTS: dict[str, str] = {
    "product_property/100393": "Cowl Neck",     # Neckline — 用户 T-shirt 套头圆领
    "product_property/100397": "Semua musim",    # Season — 全季节通用
    # 其他字段让 HiddenAttr 自动选第一个合法值
}


def _get_property_fallbacks(category: str) -> dict[str, str]:
    """Return per-category product_property fallback map.

    v1.0.10: 读每类目独立 HiddenStyle 行 + PREFERRED 默认值 + Forbid 留空。
    9 个 HiddenAttr 列对 = Template C31-C39 (C40 100403 Waist 在 HiddenAttr 但 T-shirts 无值)
    列映射:
        Template C31 (100157 Material)   → HiddenAttr C1/C2 (Katun)
        Template C32 (100198 Pattern)    → HiddenAttr C3/C4 (Polos)
        Template C33 (100393 Neckline)   → HiddenAttr C5/C6
        Template C34 (100395 Sleeve)     → HiddenAttr C7/C8
        Template C35 (100397 Season)     → HiddenAttr C9/C10
        Template C36 (100398 Style)      → HiddenAttr C11/C12
        Template C37 (100399 Fit)        → HiddenAttr C13/C14
        Template C38 (100400 Stretch)    → HiddenAttr C15/C16
        Template C39 (100401 Washing)    → 无 HiddenAttr pair → 空
        Template C40 (100403 Waist)      → HiddenAttr C17/C18
    """
    global _PROPERTY_FALLBACK_CACHE
    if _PROPERTY_FALLBACK_CACHE is not None and category in _PROPERTY_FALLBACK_CACHE:
        return _PROPERTY_FALLBACK_CACHE[category]

    from pathlib import Path

    assets_dir = Path(__file__).resolve().parent.parent / "assets"
    template_path = assets_dir / "batch-product-source.xlsx"
    if not template_path.exists():
        import sys
        if getattr(sys, "frozen", False):
            template_path = Path(sys.executable).resolve().parent / "assets" / "batch-product-source.xlsx"

    if _PROPERTY_FALLBACK_CACHE is None:
        _PROPERTY_FALLBACK_CACHE = {}

    if not template_path.exists():
        return {}

    try:
        wb = openpyxl.load_workbook(template_path, data_only=True)
        template = wb["Template"]
        hidden_style = wb["HiddenStyle"]
        hidden_attr = wb["HiddenAttr"]

        # 1. 取 Template 表头第 31-40 列的 prop_id（10 个 HiddenAttr 属性列）
        # 注：v1.0.10 官方模板 prop_id 从 C31 开始（不再有 shipping_insurance C31）
        prop_ids: list[str] = []
        for c in range(31, 41):
            v = template.cell(row=1, column=c).value
            if v:
                prop_ids.append(str(v))
            else:
                prop_ids.append(f"product_property/{100157 + (c - 31)}")

        # 2. 找每个类目在 HiddenStyle 表的独立行
        style_row_for_cat: dict[str, int] = {}
        for r in range(1, hidden_style.max_row + 1):
            cat = hidden_style.cell(row=r, column=1).value
            if cat:
                style_row_for_cat[str(cat).strip()] = r

        # 3. 对每个出现的类目，计算 fallback
        for cat, style_row in style_row_for_cat.items():
            fb: dict[str, str] = {}
            for col_idx, prop_id in enumerate(prop_ids):
                template_col = 31 + col_idx
                # v1.0.11: 用该类目自己的 HiddenStyle 行状态
                status = hidden_style.cell(row=style_row, column=template_col).value
                status = (str(status or "")).strip()
                if status == "Forbid":
                    # 印尼后台不要这个属性（不展示），留空
                    fb[prop_id] = ""
                    continue
                # v1.0.12: PREFERRED 默认值优先（覆盖 HiddenAttr 第一个值）
                preferred = _PROPERTY_PREFERRED_DEFAULTS.get(prop_id)
                if preferred:
                    fb[prop_id] = preferred
                    continue
                # v1.0.11: 用 HiddenAttr 该列对第一个合法值
                # col_idx 0-7 映射到 HiddenAttr 9 列对 (C1-C18)
                # col_idx=8 → Template C39 (100401 Care) HiddenAttr pair 8 = C17/C18
                # col_idx=9 → Template C40 (100403) 没有 HiddenAttr pair
                if col_idx > 8:
                    fb[prop_id] = ""
                    continue
                cat_col = col_idx * 2 + 1
                val_col = col_idx * 2 + 2
                if cat_col > hidden_attr.max_column or val_col > hidden_attr.max_column:
                    fb[prop_id] = ""
                    continue
                found_val = ""
                # HiddenAttr 表里 R 列是该 prop_id 在某类目的合法值列表，
                # 用 cat 列找匹配当前类目的行，取 val 列第一个非空值。
                # 注意：HiddenAttr 表有些 pair 没该类目行（如 pair 8 没 T-shirt），
                # 此时 found_val 留空，由印尼后台视为 Optional 未填。
                for r2 in range(2, hidden_attr.max_row + 1):
                    if str(hidden_attr.cell(row=r2, column=cat_col).value or "").strip() == cat:
                        v = hidden_attr.cell(row=r2, column=val_col).value
                        if v:
                            found_val = str(v).strip()
                            break
                fb[prop_id] = found_val
            _PROPERTY_FALLBACK_CACHE[cat] = fb
    except Exception as e:
        print(f"[_get_property_fallbacks] 读取模板失败: {e}", file=__import__("sys").stderr)
        return {}

    return _PROPERTY_FALLBACK_CACHE.get(category, {})


def build_rows_for_product(
    product: Product,
    settings: dict[str, Any],
    copy_suffixes: list[str] | None = None,
    apply_suffix_to_first: bool = False,
) -> list[OutputRow]:
    if copy_suffixes is None:
        copies = max(1, int(settings.get("output_copies", 1) or 1))
        copy_suffixes = [""] * copies
    else:
        copies = len(copy_suffixes) if copy_suffixes else 1

    common = _resolve_common_fields(product, settings)
    rows: list[OutputRow] = []

    if not product.variants:
        for c_idx in range(copies):
            rows.append(_build_row_for_variant(
                product, Variant(), settings, common,
                copy_idx=c_idx, copy_suffix=copy_suffixes[c_idx],
                apply_suffix_to_first=apply_suffix_to_first,
            ))
        return rows

    for v in product.variants:
        for c_idx in range(copies):
            rows.append(_build_row_for_variant(
                product, v, settings, common,
                copy_idx=c_idx, copy_suffix=copy_suffixes[c_idx],
                apply_suffix_to_first=apply_suffix_to_first,
            ))
    return rows


def write_tiktok_xlsx(
    output_path: Path,
    rows: list[OutputRow],
    template_src: Path,
) -> Path:
    """Copy the template to output_path, then write `rows` into the Template sheet."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template_src, output_path)
    wb = openpyxl.load_workbook(output_path)
    if "Template" not in wb.sheetnames:
        wb.close()
        raise ValueError(f"模板文件缺少 'Template' sheet：{template_src}")
    ws = wb["Template"]

    header_row = [(_clean_str(c.value) or "") for c in ws[1]]

    col_idx: dict[str, int] = {}
    for i, h in enumerate(header_row, 1):
        if h in TIKTOK_COLUMNS:
            col_idx[h] = i

    start_row = 2
    max_existing = ws.max_row
    if max_existing >= start_row:
        ws.delete_rows(start_row, max_existing - start_row + 1)

    for r_off, row_dict in enumerate(rows):
        excel_row = start_row + r_off
        for col_name, value in row_dict.items():
            ci = col_idx.get(col_name)
            if ci is None:
                continue
            ws.cell(row=excel_row, column=ci, value=value)

    wb.save(output_path)
    wb.close()
    return output_path
