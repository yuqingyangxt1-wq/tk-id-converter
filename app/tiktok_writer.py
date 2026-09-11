"""Write product rows into the TikTok Shop Philippines batch upload template.

The template (assets/batch-product-source.xlsx) has 12 sheets. We only write
data into the 'Template' sheet starting at row 2, leaving the rest of the
workbook (Instruction, Image, Example, HiddenStyle, HiddenAttr, etc.)
untouched so that TikTok's validator still accepts the file.

ROW STRATEGY (v3 — TikTok-compatible):
    The EasyBoss source table has one row per (color × size) SKU. We expand
    each SKU into its own TikTok listing row (each row = one variant of one
    product). This matches TikTok's official Example sheet where every
    (color, size) combination is a separate listing.

    Common fields (description, size_chart, parcel, brand, category, images,
    delivery, cod, pre_order_time) are shared across all rows of the same
    product.

    `output_copies` (防查重) duplicates each variant N times with a unique
    random suffix in the title and seller_sku, so duplicate listings can be
    detected and removed later.

    When `fill_sizes_enabled` is on, the writer falls back to the standard
    sizes list (S,M,L,XL,2XL,3XL) ONLY when the source product has no
    variation-2 values (single-SKU products). Otherwise each row keeps its
    own size value, and the field gets exactly one size per row.
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


# v1.0.0: ID 模板要求英文类目名（卖家中心 UI 也是英文）。
# EasyBoss ID 源表填的是印尼语路径（如 "Pakaian & Pakaian Dalam Pria>Atasan Pria>T-shirt"），
# 这里映射成 ID 模板/HiddenStyle 要求的英文路径。
# 完整覆盖 ID 模板 Category sheet 的 25 个男装类目。
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


# v1.0.7: 印尼 TikTok Shop 后台要求变体名（property_name_1/2）必须用印尼语，
# 同时 property_value_1 颜色值保留印尼语原值（Putih/Hitam 等）—— 不再英译。
# 之前 v1.0.0 错误地把印尼语颜色翻译成英文（Putih→White），导致后台校验报错。
# 印尼 HiddenAttr 表里 prop_id 标签虽然是英文，但用户上传的变体名/值是用户自填，
# 后台会用印尼语校验这些字段。
VAR_NAME_TRANSLATIONS: dict[str, str] = {
    # 中文 → 印尼语
    "颜色": "Warna",
    "顏色": "Warna",
    "尺码": "Ukuran",
    "尺碼": "Ukuran",
    "尺寸": "Ukuran",
    "规格": "Spesifikasi",
    # 英文 → 印尼语
    "Color": "Warna",
    "color": "Warna",
    "colour": "Warna",
    "Size": "Ukuran",
    "size": "Ukuran",
}


def _translate_var_name(name: str) -> str:
    """Translate a variation name to Indonesian for the ID template.

    Source may have Chinese ('颜色'/'尺码') or English ('Color'/'Size');
    both are mapped to Indonesian equivalents ('Warna'/'Ukuran').
    Unknown names are passed through unchanged so the user can fix them
    manually if needed.
    """
    if not name:
        return ""
    key = _clean_str(name)
    if key in VAR_NAME_TRANSLATIONS:
        return VAR_NAME_TRANSLATIONS[key]
    # Unknown name: pass through (no English fallback anymore — we want Indonesian)
    return key


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
    """Coerce a value to int/float so Excel stores it as a number.

    EasyBoss sometimes stores numeric columns as text (e.g. '356' instead of
    356). TikTok's validator and the import pipeline require numeric types
    for price / quantity / parcel dimensions / weight.
    """
    if v is None or v == "":
        return None
    if isinstance(v, bool):  # bool is a subclass of int — treat False/True as 0/1
        return int(v)
    if isinstance(v, (int, float)):
        return v
    s = str(v).strip()
    if not s:
        return None
    # Strip common currency / unit markers
    for ch in [",", " ", "₱", "$", "￥", "¥", "PHP", "php", "kg", "KG", "g", "G", "cm", "CM"]:
        s = s.replace(ch, "")
    try:
        f = float(s)
        if f.is_integer():
            return int(f)
        return f
    except ValueError:
        return v  # leave as-is


# A single output row: a dict {col_name: value}
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
        else None  # leave to per-variant price if not configured
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
    # v1.0.0: ID 版 — 把印尼语类目映射成英文（ID 模板/HiddenStyle 都用英文）
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
    """Build a single TikTok row for one (variant, copy) combination.

    ``apply_suffix_to_first=False`` (default) keeps a clean first copy (no
    suffix appended to the title or seller_sku when ``copy_idx`` is 0)
    — this matches the original "first copy is the canonical one"
    convention used for in-file duplicates.

    ``apply_suffix_to_first=True`` forces the suffix onto every row, which
    is the right behaviour for the split-output mode where each file is
    its own batch and should carry the file-wide fingerprint on every
    variant row.
    """
    use_suffix = (copy_idx > 0) or apply_suffix_to_first

    # Title
    title_prefix = settings["title_prefix"] if settings.get("title_prefix_enabled") else ""
    suffix = copy_suffix if use_suffix else ""
    title = f"{title_prefix}{product.product_name}{suffix}".strip()
    # v1.0.2: 印尼后台标题 ≤ ~100 字符更稳；源表产品名常 200+ 字符，截短
    if len(title) > 100:
        title = title[:97].rstrip() + "..."

    # Variant-level values
    var1_name = _translate_var_name(product.var1_name or "颜色")
    # v1.0.7: 印尼后台 property_value_1 接受印尼语原值（Putih/Hitam），
    # 不再做印尼语→英文翻译（之前 v1.0.0 的 translate_color 是错的，会导致后台红框）。
    var1_value = _clean_str(variant.var1)
    var2_name = _translate_var_name(product.var2_name or "尺码")

    if variant.var2:
        var2_value = _clean_str(variant.var2)
    elif settings.get("fill_sizes_enabled"):
        # Fallback: source has no size; use the standard sizes list
        std = settings.get("standard_sizes") or "S,M,L,XL,2XL,3XL"
        var2_value = std
    else:
        var2_value = ""

    # Images: prefer variant's SKU image for the main; fall back to product images
    var1_image = _clean_str(variant.sku_image)
    images = list(product.images[:9])
    while len(images) < 9:
        images.append("")
    images = [_clean_str(i) for i in images]

    # Price: prefer per-variant price, then override, then None
    price_value = _to_number(variant.price) if variant.price not in (None, "") else _to_number(common["price_override"])

    # Quantity: prefer per-variant stock, then override
    quantity_value = (
        _to_number(variant.stock) if variant.stock not in (None, "") else _to_number(common["quantity"])
    )

    # Seller SKU: use full platform_sku; if empty (e.g. EasyBoss ID source
    # 不导出 SKU 列), fall back to a stable synthetic sku so each listing
    # still has a unique non-empty identifier.
    # v1.0.1: ID 源表"平台SKU"列常为 None,直接拼接 var1+var2 生成可读 SKU
    # v1.0.2: 用 6 字符产品 hash + var1+var2 替代长产品名 slug
    #         (印尼源表产品名 200+ 字符，之前 [:60] 截断会丢掉颜色和尺码，
    #          导致同产品不同变体 seller_sku 全部重复)
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
    # main_image: prefer variant SKU image, else first product image
    row["main_image"] = var1_image if var1_image else images[0] if images else ""
    for i, img in enumerate(images[:9]):
        col = "main_image" if i == 0 else f"image_{i + 1}"
        if i == 0 and var1_image:
            # already set above
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
    # v1.0.0: ID 模板新增 2 列
    row["minimum_order_quantity"] = settings.get("minimum_order_quantity_value", 1) or 1
    row["shipping_insurance"] = (
        settings.get("shipping_insurance_value", "Optional")
        if settings.get("shipping_insurance_enabled", True)
        else ""
    )
    # v1.0.5: 给 HiddenAttr 属性 ID 列填入按 HiddenStyle/HiddenAttr 表动态计算的兜底值
    # - HiddenStyle 表某列是 "Forbid" → 留空（填了就被后台拒）
    # - HiddenAttr 表该类目下没有合法值 → 留空
    # - 否则填 HiddenAttr 表该类目的第一个合法值
    category = common.get("category", "")
    for prop_id, default_val in _get_property_fallbacks(category).items():
        row[prop_id] = default_val
    return row


# v1.0.5: 根据 HiddenStyle + HiddenAttr 表动态计算每个类目的 product_property/* 列兜底值
# v1.0.7: 修正逻辑——不再死守 HiddenStyle 表的 "Forbid" 标记。
# 实测：印尼 TikTok Shop 后台对 HiddenStyle 标 "Forbid" 的列（如 100400 Care）
# 仍然要求必填，留空被红框报 "Instruksi Mencuci" 错误。
# 安全策略：优先填 HiddenAttr 表里的第一个合法值；只在 HiddenAttr 表里没有该类目
# 适用值时才留空。这样既保证字段非空，又避免硬塞不相关类目的值。
# （印尼 HiddenStyle 表可能是国际版老模板，与印尼后台校验逻辑不同步）
_PROPERTY_FALLBACK_CACHE: dict[str, dict[str, str]] | None = None


# v1.0.6: HiddenAttr 表里的英文值映射到印尼语值
# 印尼 TikTok Shop 后台对某些字段（如季节）只接受印尼语值，拒绝英文。
# 例如 "Spring" → "Musim semi"，"All seasons" → "Semua musim"
# 注意：映射是 prop_id 维度的——只有部分字段需要翻译（不是全部）
_EN_TO_ID_VALUE_MAP: dict[str, dict[str, str]] = {
    # 100397 季节（Season）— 印尼站只接受印尼语
    "product_property/100397": {
        "Spring": "Musim semi",
        "Summer": "Musim panas",
        "Autumn": "Musim gugur",
        "Winter": "Musim dingin",
        "All seasons": "Semua musim",
    },
    # 如果以后发现其他字段需要本地化，在这里加
}


def _get_property_fallbacks(category: str) -> dict[str, str]:
    """Return per-category product_property fallback map.

    Reads the bundled template once and caches the result.
    Keys are prop_ids like 'product_property/100157'; values are HiddenAttr legal
    values, or '' (empty string) when HiddenAttr has no legal values for this category.

    v1.0.7: 不再因 HiddenStyle 标 Forbid 而留空——印尼后台实际需要这些字段填值。
    """
    global _PROPERTY_FALLBACK_CACHE
    if _PROPERTY_FALLBACK_CACHE is not None and category in _PROPERTY_FALLBACK_CACHE:
        return _PROPERTY_FALLBACK_CACHE[category]

    from pathlib import Path
    import openpyxl

    assets_dir = Path(__file__).resolve().parent.parent / "assets"
    template_path = assets_dir / "batch-product-source.xlsx"
    if not template_path.exists():
        # 在 PyInstaller 打包后，assets 应该在 exe 同目录
        import sys
        if getattr(sys, "frozen", False):
            template_path = Path(sys.executable).resolve().parent / "assets" / "batch-product-source.xlsx"

    if _PROPERTY_FALLBACK_CACHE is None:
        _PROPERTY_FALLBACK_CACHE = {}

    if not template_path.exists():
        # 没找到模板时返回空 map（保持原行为 = 不填）
        return {}

    try:
        wb = openpyxl.load_workbook(template_path, data_only=True)
        template = wb["Template"]
        hidden_style = wb["HiddenStyle"]
        hidden_attr = wb["HiddenAttr"]

        # 1. 取 Template 表头第 32-41 列的 prop_id（10 个 HiddenAttr 属性列）
        prop_ids: list[str] = []
        for c in range(32, 42):
            v = template.cell(row=1, column=c).value
            if v:
                prop_ids.append(str(v))
            else:
                # 表头空：用列号推导
                prop_ids.append(f"product_property/{100157 + (c - 32)}")

        # 2. 对 HiddenStyle 表的每个类目行（R2-R25）
        for r in range(2, hidden_style.max_row + 1):
            cat = hidden_style.cell(row=r, column=1).value
            if not cat:
                continue
            cat = str(cat).strip()
            fb: dict[str, str] = {}
            for col_idx, prop_id in enumerate(prop_ids):
                template_col = 32 + col_idx
                # v1.0.7: 忽略 HiddenStyle 状态——印尼后台校验与模板 HiddenStyle 不同步
                # Optional / Mandatory / Forbid 统一处理：填 HiddenAttr 第一个合法值
                cat_col = col_idx * 2 + 1
                val_col = col_idx * 2 + 2
                if cat_col > hidden_attr.max_column or val_col > hidden_attr.max_column:
                    fb[prop_id] = ""
                    continue
                found_val = ""
                for r2 in range(1, hidden_attr.max_row + 1):
                    if str(hidden_attr.cell(row=r2, column=cat_col).value or "").strip() == cat:
                        v = hidden_attr.cell(row=r2, column=val_col).value
                        if v:
                            found_val = str(v).strip()
                            break
                # v1.0.6: HiddenAttr 表里的英文值映射到印尼语值
                if found_val and prop_id in _EN_TO_ID_VALUE_MAP:
                    found_val = _EN_TO_ID_VALUE_MAP[prop_id].get(found_val, found_val)
                fb[prop_id] = found_val
            _PROPERTY_FALLBACK_CACHE[cat] = fb
    except Exception as e:
        # 读模板失败时返回空 map（避免阻塞转换）
        print(f"[_get_property_fallbacks] 读取模板失败: {e}", file=__import__("sys").stderr)
        return {}

    return _PROPERTY_FALLBACK_CACHE.get(category, {})


def build_rows_for_product(
    product: Product,
    settings: dict[str, Any],
    copy_suffixes: list[str] | None = None,
    apply_suffix_to_first: bool = False,
) -> list[OutputRow]:
    """Build all TikTok rows for one Product.

    Each variant becomes one row; if ``copy_suffixes`` has N entries, each
    variant is duplicated N times — once per suffix. The first suffix
    typically corresponds to copy index 0 (the "no suffix" canonical copy).

    Pass an explicit ``copy_suffixes`` of any length to control duplication;
    the per-product ``output_copies`` setting is only used as a fallback
    when ``copy_suffixes`` is None.

    ``apply_suffix_to_first`` defaults to False (i.e. the first copy has a
    clean title and seller_sku). Set True for split-file mode where the
    whole file should carry a single random fingerprint.
    """
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

    # v1.0.1: TikTok 卖家中心官方要求 "Don't add or delete any rows or columns"
    # — 之前 v1.0.0 通过 delete_cols 删除 pre_order_time 会让印尼/泰国/菲律宾
    # 后台识别模板不完整、整批 20 个产品被拒。现改为保留 pre_order_time 列
    # 并把它的值设为空（无预售时后台允许空）。
    header_row = [(_clean_str(c.value) or "") for c in ws[1]]

    # Build header→col index from the first row of the Template sheet
    col_idx: dict[str, int] = {}
    for i, h in enumerate(header_row, 1):
        if h in TIKTOK_COLUMNS:
            col_idx[h] = i

    # Find first data row (the template's data area usually starts at row 2;
    # skip any pre-existing instruction rows by locating the header row)
    start_row = 2
    # Clear any pre-existing data rows in the Template sheet
    max_existing = ws.max_row
    if max_existing >= start_row:
        ws.delete_rows(start_row, max_existing - start_row + 1)

    # Write new rows
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