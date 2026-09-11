"""Configuration management - load/save config.json, provide defaults."""
from __future__ import annotations

import json
import os
import sys
from pathlib import Path
from typing import Any


APP_NAME = "TK印尼表格转化工具"
__version__ = "1.0.10"  # ID v1.0.10: 用 TikTok Shop 官方 ID 模板（40列全印尼语）+ 删 shipping_insurance 列 + 印尼语类目/品牌/属性值


def get_app_dir() -> Path:
    """Return the writable directory next to the executable.

    This is where config.json (and any future user files) live. For both
    a PyInstaller --onefile exe and a `python -m app.main` source run,
    this is the directory containing the exe / project root.
    """
    if getattr(sys, "frozen", False):  # PyInstaller bundle
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def get_assets_dir() -> Path:
    """Return the directory containing bundled read-only assets (the template).

    Source run: <project>/assets/
    PyInstaller: also <exe-dir>/assets/ (we ship assets/ next to the exe
                 rather than relying on _MEIPASS, matching the original
                 tool's "template/ folder" layout and keeping the template
                 user-replaceable).
    """
    return get_app_dir() / "assets"


def default_config() -> dict[str, Any]:
    """The shipped defaults — mirror the original v1.2 tool."""
    return {
        "version": 2,
        "product_xlsx_last": "",
        "product_xlsx_output_dir": "",
        "product_xlsx_last_output_dir": "",
        "product_pool_dir": "",
        "product_xlsx_settings": {
            # v1.0.0: 让用户能跳过图片下载（默认勾上 = 下载；GUI 转化页有开关）
            "download_images_enabled": True,
            # Title (ID 版：Rp 卢比默认)
            "title_prefix_enabled": True,
            "title_prefix": "Kaos Unisex Oversize ",
            # Brand
            "brand_enabled": True,
            "brand_value": "Tidak ada merek",
            # Price (IDR Rp)
            "price_enabled": True,
            "price_value": 115000,
            # Quantity
            "quantity_enabled": True,
            "quantity_value": 999,
            # COD (v1.0.9: 印尼市场支持 COD = Y，参考 EasyBoss 输出用 Y)
            "cod_enabled": True,
            "cod_value": "Y",
            # Fill standard sizes S-3XL into property_value_2
            "fill_sizes_enabled": True,
            "standard_sizes": "S,M,L,XL,2XL,3XL",
            # Random suffix to differentiate copies
            "title_random_suffix_enabled": True,
            "random_suffix_length": 3,
            # Category (ID 模板 25 个男装类目，默认 T-shirt)
            "category_enabled": True,
            "category_value": "Atasan Pria/T-shirt",
            # Output copies per product
            "output_copies": 2,
            "split_output_files": False,
            # Description (ID 市场英文通用)
            "description_enabled": True,
            "description_value": (
                "Premium quality product with careful packaging. "
                "Material: soft and comfortable fabric, breathable for daily wear. "
                "Care: machine washable, retains shape after washing. "
                "Size: please refer to the size chart image before ordering. "
                "Pengiriman: pesanan dikirim 1-2 hari kerja; estimasi sampai 3-8 hari."
            ),
            # v1.0.9: Size chart 默认空——印尼后台只接受 Media Center URL（https://p16-oec-sg.ibyteimg.com/...），
            # GitHub raw URL 会失败。用户需要手动上传图片到 TikTok Shop 后台 Media Center，
            # 然后把返回的 URL 填到这里（GUI 有输入框）。
            "size_chart_enabled": True,
            "size_chart_value": "",
            # Parcel
            "parcel_enabled": True,
            "parcel_weight_value": 200,   # grams
            "parcel_length_value": 10,   # cm
            "parcel_width_value": 10,
            "parcel_height_value": 5,
            # ID-specific fields
            "minimum_order_quantity_value": 1,    # C28, 最低购买量（1-20 整数）
            "shipping_insurance_enabled": True,
            "shipping_insurance_value": "Optional",  # ID 后台用字面量 "Optional"
            # TikTok-specific
            "pre_order_time_value": "",  # 留空 = 无预售（ID 模板允许）
            "delivery_value": "",   # ID 模板 delivery 列空 = 跟店铺默认（印尼模板不强制）
        },
        # Source column mapping (override which EasyBoss/源列 maps to what)
        # If a key is empty/None, the reader falls back to auto-detection.
        "source_column_mapping": {
            "product_name": "产品名",
            "brand": "品牌",
            "category": "产品类目",
            "platform_sku": "平台SKU",
            "var1_name": "规格1名称",
            "var1_value": "规格1选项",
            "var2_name": "规格2名称",
            "var2_value": "规格2选项",
            "var3_name": "规格3名称",
            "var3_value": "规格3选项",
            "price": "税前价格",
            "stock": "库存",
            "sku_image": "SKU图片",
            "image_1": "产品图片1",
            "image_2": "产品图片2",
            "image_3": "产品图片3",
            "image_4": "产品图片4",
            "image_5": "产品图片5",
            "image_6": "产品图片6",
            "image_7": "产品图片7",
            "image_8": "产品图片8",
            "image_9": "产品图片9",
            "parcel_weight_kg": "包裹重量（KG）",
            "parcel_length": "包裹长度（CM）",
            "parcel_width": "包裹宽度（CM）",
            "parcel_height": "包裹高度（CM）",
            "description": "产品描述",
            "size_chart": "尺码图",
        },
    }


def config_path() -> Path:
    return get_app_dir() / "config.json"


def load_config(path: Path | None = None) -> dict[str, Any]:
    """Load config from disk, merging with defaults so new keys appear.

    v1.0.9: 不再强制改 cod_value（之前 v1.0.7 的"Y→N"自动重置是错的——
    印尼市场参考 EasyBoss 输出 cod="Y"）。让 config.json 完全控制。
    """
    p = path or config_path()
    cfg = default_config()
    if p.exists():
        try:
            with p.open("r", encoding="utf-8") as f:
                on_disk = json.load(f)
            # Shallow merge at top level, deep merge for *_settings / mapping
            for k, v in on_disk.items():
                if isinstance(v, dict) and isinstance(cfg.get(k), dict):
                    cfg[k].update(v)
                else:
                    cfg[k] = v
        except (OSError, json.JSONDecodeError):
            # Corrupt config → fall back to defaults but don't overwrite
            pass

    return cfg


def save_config(cfg: dict[str, Any], path: Path | None = None) -> None:
    p = path or config_path()
    p.parent.mkdir(parents=True, exist_ok=True)
    with p.open("w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def get_settings(cfg: dict[str, Any]) -> dict[str, Any]:
    """Return the product_xlsx_settings sub-dict, applying defaults for missing keys."""
    defaults = default_config()["product_xlsx_settings"]
    s = dict(defaults)
    s.update(cfg.get("product_xlsx_settings", {}))
    return s


def update_settings(cfg: dict[str, Any], **kwargs: Any) -> None:
    cfg.setdefault("product_xlsx_settings", {})
    cfg["product_xlsx_settings"].update(kwargs)
