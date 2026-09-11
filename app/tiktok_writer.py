"""Write product rows into the TikTok Shop Indonesia batch upload template.

The template (assets/batch-product-source.xlsx) has 12 sheets. We only write
data into the 'Template' sheet starting at row 2, leaving the rest of the
workbook (Instruction, Image, Example, HiddenStyle, HiddenAttr, etc.)
untouched so that TikTok's validator still accepts the file.

ROW STRATEGY (v1.0.9 — based on reference EasyBoss output):
    The EasyBoss source table has one row per (color × size) SKU. We expand
    each SKU into its own TikTok listing row (each row = one variant of one
    product).

    v1.0.9 — BREAKING CHANGE (终于修对了印尼隐藏属性列):
    1. 不再翻译印尼语颜色 → 英文 (Putih/Hitam 直接透传，参考工具也是用印尼语)
    2. property_name_1/2 保持英文 "Color" / "Size"
    3. product_property/* 字段改用参考模板的 HiddenAttr 真实合法值:
       - 100397 Season: "All seasons" (EasyBoss 默认)
       - 100398 Style: "Basic"
       - 100399 Fit: "Fitted"
       - 100400 Stretch: 空 (Forbid for T-shirts)
       - 100401 Washing: 空 (T-shirts 在 HiddenAttr 无合法值)
       - 100403 Waist: 空 (Forbid)
    4. HiddenStyle 表按每个类目独立行读，不是用 R1 header (R1 是 Waistcoats & Gilets)
    5. seller_sku 用印尼语颜色名 (YW001-PUTIH-S) 而不是英文
    6. cod 默认 Y (印尼后台接受，与 EasyBoss 一致)
    7. size_chart 不默认填 GitHub raw URL (印尼后台只接受 Media Center URL)
"""
from __future__ import annotations

import shutil
from pathlib import Path
from typing import Any

import openpyxl

from .config import get_assets_dir
from .source_reader import Product, Variant


TEMPLATE_FILENAME = "batch-product-source.xlsx"

# The 39 Template-sheet column headers, in order
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
    "minimum_order_quantity",  # v1.0.0: ID 模板新增 C28
    "size_chart",
    "cod",
    "shipping_insurance",  # v1.0.0: ID 模板新增 C31（字面量 Optional / Y / N）
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


# v1.0.9: ID 模板要求英文类目名（卖家中心 UI 也是英文）。
# EasyBoss ID 源表填的是印尼语路径（如 "Pakaian & Pakaian Dalam Pria>Atasan Pria>T-shirt"），
# 这里映射成 ID 模板/HiddenStyle 要求的英文路径。
_ID_CATEGORY_TRANSLATION: dict[str, str] = {
    # 男装上衣
    "Pakaian & Pakaian Dalam Pria>Atasan Pria>T-shirt": "Men's Tops/T-shirts",
    "Pakaian & Pakaian Dalam Pria>Atasan Pria>Kaus Polo": "Men's Tops/Polo Shirts",
    "Pakaian & Pakaian Dalam Pria>Atasan Pria>Kemeja": "Men's Tops/Shirts",
    "Pakaian & Pakaian Dalam Pria>Atasan Pria>Rajut": "Men's Tops/Knitwear",
    "Pakaian & Pakaian Dalam Pria>Atasan Pria>Hoodie & Sweatshirt": "Men's Tops/Hoodies & Sweatshirts",
    "Pakaian & Pakaian Dalam Pria>Atasan Pria>Rompi & Gilet": "Men's Tops/Waistcoats & Gilets",
    "Pakaian & Pakaian Dalam Pria>Atasan Pria>Jaket & Mantel": "Men's Tops/Jackets & Coats",
    # 男装下装
    "Pakaian & Pakaian Dalam Pria>Celana Pria>Celana Pendek": "Men's Bottoms/Men's Shorts",
    "Pakaian & Pakaian Dalam Pria>Celana Pria>Jeans": "Men's Bottoms/Men's Jeans",
    "Pakaian & Pakaian Dalam Pria>Celana Pria>Celana Panjang": "Men's Bottoms/Men's Pants",
    # 男士套装
    "Setelan Pria>Set Pakaian Pria": "Men's Suits & Sets/Men's Clothing Sets",
    "Setelan Pria>Setelan Pria": "Men's Suits & Sets/Men's Suits",
    "Setelan Pria>Overall": "Men's Suits & Sets/Overalls",
    # 男士内衣袜子
    "Pakaian Dalam & Kaus Kaki Pria>Pakaian Dalam Pria": "Men's Underwear & Socks/Men's Underwear",
    "Pakaian Dalam & Kaus Kaki Pria>Tank Top & Pakaian Dalam": "Men's Underwear & Socks/Men's Tanks & Undershirts",
    "Pakaian Dalam & Kaus Kaki Pria>Pakaian Dalam Hangat": "Men's Underwear & Socks/Men's Thermal Underwear",
    "Pakaian Dalam & Kaus Kaki Pria>Kaus Kaki": "Men's Underwear & Socks/Socks",
    # 男士睡衣
    "Pakaian Tidur & Pakaian Santai Pria>Piyama & Loungewear": "Men's Sleepwear & Loungewear/Pajamas & Loungewear",
    "Pakaian Tidur & Pakaian Santai Pria>Robe Pria": "Men's Sleepwear & Loungewear/Men's Robes",
    "Pakaian Tidur & Pakaian Santai Pria>Nightshirt": "Men's Sleepwear & Loungewear/Nightshirts",
    "Pakaian Tidur & Pakaian Santai Pria>Piyama Pria": "Men's Sleepwear & Loungewear/Men's Pajamas",
    # 男士特殊场合
    "Pakaian Acara Khusus Pria>Kostum": "Men's Special Occasion Clothing/Costumes",
    "Pakaian Acara Khusus Pria>Pakaian Kerja": "Men's Special Occasion Clothing/Workwear",
    "Pakaian Acara Khusus Pria>Pakaian Tradisional": "Men's Special Occasion Clothing/Traditional Wear",
}


def translate_category(raw: str) -> str:
    """Translate an EasyBoss ID Indonesian category path into the English
    category path that ID  template + HiddenStyle + HiddenAttr require.

    Returns the original string if no mapping is found (so that the user can
    see what was untranslated in the output and fix it manually).
    """
    if not raw:
        return raw
    return _ID_CATEGORY_TRANSLATION.get(raw, raw)


# v1.0.9: 变体名保持英文 "Color" / "Size"，property_value_* 保留印尼语原值（Putih/Hitam）
# 参考 EasyBoss 输出就是这种混合模式。
VAR_NAME_TRANSLATIONS: dict[str, str] = {
    # Chinese → English
    "颜色": "Color",
    "顏色": "Color",
    "尺码": "Size",
    "尺碼": "Size",
    "尺寸": "Size",
    "规格": "Specification",
    # English pass-through (normalize capitalization)
    "Color": "Color",
    "color": "Color",
    "colour": "Color",
    "Size": "Size",
    "size": "Size",
}


def _translate_var_name(name: str) -> str:
    """Translate a variation name to English for the ID template.

    v1.0.9: 保持英文（参考 EasyBoss 输出）。
    """
    if not name:
        return ""
    key = _clean_str(name)
    if key in VAR_NAME_TRANSLATIONS:
        return VAR_NAME_TRANSLATIONS[key]
    return key


# v1.0.9 不再翻译颜色。源表的印尼语颜色（Putih/Hitam）直接透传到 property_value_1


def template_path() -> Path:
    """Locate the template inside the app's assets/ directory."""
    return get_assets_dir() / TEMPLATE_FILENAME


def _kg_to_grams(v: Any) -> Any:
    """EasyBoss source stores weight in KG; TikTok wants grams."""
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
    """Coerce a value to int/float so Excel stores it as a number."""
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
    brand = (
        settings["brand_value"]
        if settings.get("brand_enabled") and settings.get("brand_value")
        else (product.brand or "")
    )
    if not brand:
        brand = "No brand"

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

    # Title
    title_prefix = settings["title_prefix"] if settings.get("title_prefix_enabled") else ""
    suffix = copy_suffix if use_suffix else ""
    title = f"{title_prefix}{product.product_name}{suffix}".strip()
    # v1.0.8: 印尼后台不接受 en dash / em dash 字符（U+2013, U+2014）
    title = (
        title.replace("\u2013", " - ")
             .replace("\u2014", " - ")
             .replace("\u2019", "'")
             .replace("\u201c", '"')
             .replace("\u201d", '"')
    )
    if len(title) > 100:
        title = title[:97].rstrip() + "..."

    # Variant-level values
    var1_name = _translate_var_name(product.var1_name or "颜色")
    # v1.0.9: 保留印尼语颜色原值 (Putih/Hitam 直接透传，不再翻译)
    var1_value = _clean_str(variant.var1)
    var2_name = _translate_var_name(product.var2_name or "尺码")

    if variant.var2:
        var2_value = _clean_str(variant.var2)
    elif settings.get("fill_sizes_enabled"):
        std = settings.get("standard_sizes") or "S,M,L,XL,2XL,3XL"
        var2_value = std
    else:
        var2_value = ""

    # Images
    var1_image = _clean_str(variant.sku_image)
    images = list(product.images[:9])
    while len(images) < 9:
        images.append("")
    images = [_clean_str(i) for i in images]

    # Price
    price_value = _to_number(variant.price) if variant.price not in (None, "") else _to_number(common["price_override"])

    # Quantity
    quantity_value = (
        _to_number(variant.stock) if variant.stock not in (None, "") else _to_number(common["quantity"])
    )

    # Seller SKU: v1.0.9 用印尼语颜色名 (Putih/Hitam 等) 而不是英文
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
    row["shipping_insurance"] = (
        settings.get("shipping_insurance_value", "Optional")
        if settings.get("shipping_insurance_enabled", True)
        else ""
    )
    # v1.0.9: 读每类目独立行的 HiddenStyle 状态 + PREFERRED 默认值
    category = common.get("category", "")
    for prop_id, default_val in _get_property_fallbacks(category).items():
        row[prop_id] = default_val
    return row


# v1.0.9: 完整重写 _get_property_fallbacks
# 关键改动:
# 1. HiddenStyle 表按每个类目独立行 (R1=Waistcoats, R11=T-shirts) — 之前用 R1 header 是错的
# 2. PREFERRED 默认值 (参考 EasyBoss 输出: All seasons / Basic / Fitted)
# 3. Forbid 列 → 留空 (T-shirts R11: C39 100400 Stretch / C41 100403 Waist 都是 Forbid)
# 4. 如果 HiddenAttr 无该类目合法值 → 留空 (T-shirts C40 100401 Washing = 空)
_PROPERTY_FALLBACK_CACHE: dict[str, dict[str, str]] | None = None


# v1.0.9: PREFERRED 默认值 — 参考印尼 EasyBoss 工具的输出（更通用的安全值）
_PROPERTY_PREFERRED_DEFAULTS: dict[str, str] = {
    "product_property/100397": "All seasons",   # Season
    "product_property/100398": "Basic",          # Style
    "product_property/100399": "Fitted",         # Fit
}


def _get_property_fallbacks(category: str) -> dict[str, str]:
    """Return per-category product_property fallback map.

    v1.0.9: 读每类目独立 HiddenStyle 行 + PREFERRED 默认值 + Forbid 留空。
    列映射:
        Template C32 (100157 Material)  → HiddenAttr C1/C2
        Template C33 (100198 Pattern)   → HiddenAttr C3/C4
        Template C34 (100393 Neckline)  → HiddenAttr C5/C6
        Template C35 (100395 Sleeve)    → HiddenAttr C7/C8
        Template C36 (100397 Season)    → HiddenAttr C9/C10
        Template C37 (100398 Style)     → HiddenAttr C11/C12
        Template C38 (100399 Fit)       → HiddenAttr C13/C14
        Template C39 (100400 Stretch)   → HiddenAttr C15/C16
        Template C40 (100401 Washing)   → HiddenAttr C17/C18
        Template C41 (100403 Waist)     → 无 HiddenAttr 列对 → 始终空
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

        # 1. 取 Template 表头第 32-41 列的 prop_id
        prop_ids: list[str] = []
        for c in range(32, 42):
            v = template.cell(row=1, column=c).value
            if v:
                prop_ids.append(str(v))
            else:
                prop_ids.append(f"product_property/{100157 + (c - 32)}")

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
                template_col = 32 + col_idx
                # v1.0.9: 用该类目自己的 HiddenStyle 行状态 (不是 R1 header!)
                status = hidden_style.cell(row=style_row, column=template_col).value
                status = (str(status or "")).strip()
                if status == "Forbid":
                    fb[prop_id] = ""
                    continue
                # 优先用 PREFERRED 默认值
                preferred = _PROPERTY_PREFERRED_DEFAULTS.get(prop_id)
                if preferred:
                    fb[prop_id] = preferred
                    continue
                # 用 HiddenAttr 该列对第一个合法值 (按当前类目过滤)
                if col_idx >= 9:
                    # 第 10 列 (100403 Waist) 没有 HiddenAttr 列对 → 空
                    fb[prop_id] = ""
                    continue
                cat_col = col_idx * 2 + 1
                val_col = col_idx * 2 + 2
                if cat_col > hidden_attr.max_column or val_col > hidden_attr.max_column:
                    fb[prop_id] = ""
                    continue
                found_val = ""
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
