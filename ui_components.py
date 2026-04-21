from __future__ import annotations

import streamlit as st

from address_utils import normalize_text, split_address


def render_product_row(products: list[dict[str, str]], key_prefix: str, row_number: int) -> dict[str, object]:
    product_names = [product["name"] for product in products]
    row_cols = st.columns([0.8, 5, 2, 1])

    with row_cols[0]:
        st.caption(f"第 {row_number} 行")

    with row_cols[1]:
        selected_product = st.selectbox(
            f"产品 {key_prefix}",
            product_names,
            key=f"product_{key_prefix}",
            label_visibility="collapsed",
            index=None,
            placeholder="搜索并选择产品",
        )
        selected_sku = next(
            (product["sku_code"] for product in products if product["name"] == selected_product),
            "",
        )

    with row_cols[2]:
        st.text_input(
            f"产品六九码 {key_prefix}",
            value=selected_sku or "未匹配到69码",
            disabled=True,
            label_visibility="collapsed",
        )

    with row_cols[3]:
        quantity = st.number_input(
            f"产品数量 {key_prefix}",
            min_value=1,
            step=1,
            key=f"quantity_{key_prefix}",
            label_visibility="collapsed",
        )

    return {
        "product_name": selected_product,
        "sku_code": selected_sku,
        "quantity": int(quantity),
    }


def render_product_table(
    products: list[dict[str, str]],
    key_prefix: str,
    row_count: int,
    title: str,
) -> list[dict[str, object]]:
    with st.container(border=True):
        st.markdown(f"#### {title}")
        header_cols = st.columns([0.8, 5, 2, 1])
        header_cols[0].caption("行号")
        header_cols[1].caption("产品")
        header_cols[2].caption("产品六九码")
        header_cols[3].caption("数量")

        items: list[dict[str, object]] = []
        for item_index in range(row_count):
            items.append(
                render_product_row(
                    products=products,
                    key_prefix=f"{key_prefix}_{item_index}",
                    row_number=item_index + 1,
                )
            )
        return items


def render_address_preview(address_text: str, address_index: int) -> None:
    normalized = normalize_text(address_text)
    if not normalized:
        st.caption("填写地址后，这里会显示解析预览。")
        return

    try:
        parsed = split_address(normalized)
        st.success(f"地址 {address_index + 1} 解析正常")
        preview_cols = st.columns(6)
        preview_cols[0].metric("收货人", parsed["receiver"])
        preview_cols[1].metric("手机", parsed["phone"])
        preview_cols[2].metric("省", parsed["province"])
        preview_cols[3].metric("市", parsed["city"])
        preview_cols[4].metric("区", parsed["district"])
        preview_cols[5].metric("详细地址", parsed["detail_address"])
    except ValueError as exc:
        st.error(f"地址 {address_index + 1} 解析失败：{exc}")
