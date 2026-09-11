"""Dryrun v1.0.7: 印尼后台本地化 + 100400 Care 填值"""
import sys
from pathlib import Path

sys.path.insert(0, '/Users/apple/WorkBuddy/2026-09-08-17-31-48/tk-id-converter-repo')

from app.config import default_config, load_config
from app.tiktok_writer import (
    TIKTOK_COLUMNS, _translate_var_name, _get_property_fallbacks, build_rows_for_product,
    write_tiktok_xlsx,
)
from app.source_reader import read_source
from collections import Counter


def main():
    # 1. 直接测试 _translate_var_name
    print("=== _translate_var_name 测试 ===")
    cases = ['颜色', 'Color', 'color', '尺码', 'Size', 'size', '随机新名']
    for c in cases:
        print(f"  '{c}' -> '{_translate_var_name(c)}'")
    print()

    # 2. 直接测试 _get_property_fallbacks
    print("=== _get_property_fallbacks 测试 ===")
    fb = _get_property_fallbacks("Men's Tops/T-shirts")
    for k, v in fb.items():
        print(f"  {k}: {v!r}")
    print()

    # 3. 加载配置（强制 cod 重置）
    print("=== config.py 测试 ===")
    cfg = load_config()
    s = cfg.get("product_xlsx_settings", {})
    print(f"  cod_value: {s.get('cod_value')}")
    print(f"  cod_enabled: {s.get('cod_enabled')}")
    print()

    # 4. 完整 dryrun
    print("=== 完整 dryrun：240 行输出验证 ===")
    source_path = Path('/Users/apple/Desktop/导出#SKU_2026_09_11_17_15_46.xlsx')
    products = read_source(source_path)
    print(f"  读到 {len(products)} 个产品")

    all_rows = []
    for p in products:
        rows = build_rows_for_product(p, s, copy_suffixes=[""])
        all_rows.extend(rows)
    print(f"  生成 {len(all_rows)} 行")

    # 验证关键列
    print("\n=== 关键列值分布 ===")
    for col_name in ['property_name_1', 'property_name_2', 'property_value_1', 'property_value_2',
                     'delivery', 'cod', 'shipping_insurance', 'minimum_order_quantity']:
        if col_name in TIKTOK_COLUMNS:
            ci = TIKTOK_COLUMNS.index(col_name)
            vals = [r.get(col_name, '') for r in all_rows]
            c = Counter(str(v) if v else '<空>' for v in vals)
            print(f"  {col_name}: {dict(c)}")

    print("\n=== product_property 10 列值分布 ===")
    for prop_id in TIKTOK_COLUMNS[31:41]:  # product_property/100157 ~ 100403
        vals = [r.get(prop_id, '') for r in all_rows]
        c = Counter(str(v) if v else '<空>' for v in vals)
        print(f"  {prop_id}: {dict(c)}")

    # 5. 写入测试 xlsx 并 dry-run 验证 seller_sku 唯一
    print("\n=== seller_sku 唯一性 ===")
    skus = [r.get('seller_sku', '') for r in all_rows]
    print(f"  总数: {len(skus)}, 唯一: {len(set(skus))}")

    # 6. 写一个测试 xlsx
    output_path = Path('/tmp/dryrun_v107.xlsx')
    template_path = Path('/Users/apple/WorkBuddy/2026-09-08-17-31-48/tk-id-converter-repo/assets/batch-product-source.xlsx')
    write_tiktok_xlsx(output_path, all_rows, template_path)
    print(f"\n=== 写入: {output_path} ===")
    print(f"  大小: {output_path.stat().st_size} bytes")


if __name__ == '__main__':
    main()