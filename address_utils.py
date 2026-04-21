from __future__ import annotations

import re


MUNICIPALITIES = ("北京市", "上海市", "天津市", "重庆市")
PROVINCE_ALIASES = {
    "北京": "北京市",
    "上海": "上海市",
    "天津": "天津市",
    "重庆": "重庆市",
    "河北": "河北省",
    "山西": "山西省",
    "辽宁": "辽宁省",
    "吉林": "吉林省",
    "黑龙江": "黑龙江省",
    "江苏": "江苏省",
    "浙江": "浙江省",
    "安徽": "安徽省",
    "福建": "福建省",
    "江西": "江西省",
    "山东": "山东省",
    "河南": "河南省",
    "湖北": "湖北省",
    "湖南": "湖南省",
    "广东": "广东省",
    "海南": "海南省",
    "四川": "四川省",
    "贵州": "贵州省",
    "云南": "云南省",
    "陕西": "陕西省",
    "甘肃": "甘肃省",
    "青海": "青海省",
    "台湾": "台湾省",
    "广西": "广西壮族自治区",
    "内蒙古": "内蒙古自治区",
    "西藏": "西藏自治区",
    "宁夏": "宁夏回族自治区",
    "新疆": "新疆维吾尔自治区",
    "香港": "香港特别行政区",
    "澳门": "澳门特别行政区",
}
CITY_TO_PROVINCE = {
    "广州市": "广东省",
    "深圳市": "广东省",
    "佛山市": "广东省",
    "东莞市": "广东省",
    "中山市": "广东省",
    "珠海市": "广东省",
    "惠州市": "广东省",
    "茂名市": "广东省",
    "北海市": "广西壮族自治区",
    "南宁市": "广西壮族自治区",
    "桂林市": "广西壮族自治区",
    "海口市": "海南省",
    "三亚市": "海南省",
}
CITY_NAME_ALIASES = {
    "张家口": "张家口市",
    "廊坊": "廊坊市",
    "太原": "太原市",
    "郑州": "郑州市",
    "茂名": "茂名市",
    "西安": "西安市",
    "济南": "济南市",
    "大连": "大连市",
    "长沙": "长沙市",
    "北海": "北海市",
    "广州": "广州市",
}
ADDRESS_START_PATTERN = re.compile(
    r"(北京市|上海市|天津市|重庆市|内蒙古自治区|广西壮族自治区|西藏自治区|宁夏回族自治区|新疆维吾尔自治区|香港特别行政区|澳门特别行政区|[^,，、；;\s]{2,}省|[^,，、；;\s]{2,}自治区)"
)


def normalize_text(value: object) -> str:
    return str(value).strip() if value is not None else ""


def normalize_address_text(raw_address: str) -> str:
    text = str(raw_address).strip()
    text = re.sub(r"[：:]", " ", text)
    text = re.sub(r"[，,、；;|/]+", " ", text)
    text = re.sub(r"([^\W\d_])\s*(1\d{10})", r"\1 \2", text)
    text = re.sub(r"(1\d{10})\s*([^\W\d_])", r"\1 \2", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def find_address_start(text: str) -> int:
    match = ADDRESS_START_PATTERN.search(text)
    if match:
        return match.start()
    for short_name in sorted(PROVINCE_ALIASES.keys(), key=len, reverse=True):
        idx = text.find(short_name)
        if idx >= 0:
            return idx
    city_match = re.search(r"(北京市|上海市|天津市|重庆市|[^,，、；;\s]{2,}市)", text)
    if city_match:
        return city_match.start()
    for short_city in sorted(CITY_NAME_ALIASES.keys(), key=len, reverse=True):
        idx = text.find(short_city)
        if idx >= 0:
            return idx
    return -1


def is_address_token(token: str) -> bool:
    if not token:
        return False
    if ADDRESS_START_PATTERN.search(token):
        return True
    return bool(
        re.search(
            r"(省|市|区|县|旗|镇|乡|街道|路|街|道|巷|弄|号|栋|幢|单元|室|楼|层|\d)",
            token,
        )
    )


def normalize_region_aliases(address_body: str) -> str:
    normalized = address_body.strip()
    for short_name, full_name in sorted(PROVINCE_ALIASES.items(), key=lambda item: len(item[0]), reverse=True):
        if normalized.startswith(full_name):
            return normalized
        if normalized.startswith(short_name):
            remainder = normalized[len(short_name) :]
            if remainder.startswith(("省", "市", "自治区", "特别行政区")):
                return normalized
            return f"{full_name}{remainder}"
    return normalized


def split_name_and_address(text_without_phone: str) -> tuple[str, str]:
    text_without_phone = normalize_region_aliases(text_without_phone)
    address_start = find_address_start(text_without_phone)
    if address_start < 0:
        raise ValueError("地址中未识别到省、市或直辖市信息。")

    tokens = [token.strip(" ,，、；;") for token in text_without_phone.split(" ") if token.strip(" ,，、；;")]
    if len(tokens) == 1:
        raise ValueError("地址中未识别到收货人名称。")

    non_address_tokens = [token for token in tokens if not is_address_token(token)]
    if len(non_address_tokens) == 1:
        receiver = non_address_tokens[0]
        removed = False
        rebuilt_tokens = []
        for token in tokens:
            if token == receiver and not removed:
                removed = True
                continue
            rebuilt_tokens.append(token)
        address_candidate = "".join(rebuilt_tokens)
        return receiver, address_candidate

    if not is_address_token(tokens[0]) and any(is_address_token(token) for token in tokens[1:]):
        return tokens[0], "".join(tokens[1:])

    if not is_address_token(tokens[-1]) and any(is_address_token(token) for token in tokens[:-1]):
        return tokens[-1], "".join(tokens[:-1])

    prefix = text_without_phone[:address_start].strip(" ,，、；;")
    suffix = text_without_phone[address_start:].strip(" ,，、；;")
    if prefix:
        return prefix, suffix

    raise ValueError("地址中未识别到收货人名称。")


def split_district_and_detail(rest: str) -> tuple[str, str]:
    suffixes = ["街道", "自治县", "县", "区", "市", "旗", "镇", "乡"]
    best_end = None
    best_suffix = ""
    for suffix in suffixes:
        match = re.search(rf"^(.+?{suffix})", rest)
        if not match:
            continue
        candidate = match.group(1)
        if len(candidate) < 2:
            continue
        if best_end is None or len(candidate) < best_end:
            best_end = len(candidate)
            best_suffix = candidate

    if not best_suffix:
        raise ValueError("地址中未识别到区/县。")

    detail_address = rest[len(best_suffix) :].strip()
    if not detail_address:
        raise ValueError("地址中缺少详细地址。")
    return best_suffix, detail_address


def split_region_fields(address_body: str) -> dict[str, str]:
    if not address_body:
        raise ValueError("请输入完整地址。")
    address_body = normalize_region_aliases(address_body)

    for municipality in MUNICIPALITIES:
        if address_body.startswith(municipality):
            rest = address_body[len(municipality) :].strip()
            district, detail_address = split_district_and_detail(rest)
            return {
                "province": municipality,
                "city": municipality,
                "district": district,
                "detail_address": detail_address,
            }

    province_match = re.match(r"^([^ ]+?省|[^ ]+?自治区)", address_body)
    if province_match:
        province = province_match.group(1)
        rest = address_body[province_match.end() :].strip()
    else:
        city_fallback_match = re.match(r"^([^ ]+?市)", address_body)
        if not city_fallback_match:
            raise ValueError("地址中未识别到省份。")
        city_name = city_fallback_match.group(1)
        if city_name not in CITY_TO_PROVINCE:
            raise ValueError("地址中未识别到省份。")
        province = CITY_TO_PROVINCE[city_name]
        rest = address_body

    city_match = re.match(r"^([^ ]+?市|[^ ]+?自治州|[^ ]+?地区|[^ ]+?盟)", rest)
    if city_match:
        city = city_match.group(1)
        rest = rest[city_match.end() :].strip()
    else:
        city = ""
        for short_city, full_city in sorted(CITY_NAME_ALIASES.items(), key=lambda item: len(item[0]), reverse=True):
            if rest.startswith(full_city):
                city = full_city
                rest = rest[len(full_city) :].strip()
                break
            if rest.startswith(short_city):
                city = full_city
                rest = rest[len(short_city) :].strip()
                break
        if not city:
            raise ValueError("地址中未识别到城市。")

    district, detail_address = split_district_and_detail(rest)
    return {
        "province": province,
        "city": city,
        "district": district,
        "detail_address": detail_address,
    }


def split_address(raw_address: str) -> dict[str, str]:
    cleaned = normalize_address_text(raw_address)
    if not cleaned:
        raise ValueError("请输入完整地址。")

    phone_match = re.search(r"1\d{10}", cleaned)
    if not phone_match:
        raise ValueError("地址中未识别到 11 位手机号。")

    phone = phone_match.group()
    text_without_phone = f"{cleaned[: phone_match.start()]} {cleaned[phone_match.end() :]}"
    text_without_phone = re.sub(r"\s+", " ", text_without_phone).strip(" ,，、；;")
    if not text_without_phone:
        raise ValueError("手机号之外缺少姓名和详细地址。")

    receiver, address_body = split_name_and_address(text_without_phone)
    region_info = split_region_fields(address_body)
    return {
        "receiver": receiver,
        "phone": phone,
        **region_info,
    }


def validate_address_groups(address_groups: list[dict[str, object]]) -> list[str]:
    errors: list[str] = []
    seen_order_numbers: dict[str, str] = {}
    for address_index, address_group in enumerate(address_groups, start=1):
        address_text = normalize_text(address_group["address_text"])
        if not address_text:
            continue
        try:
            split_address(address_text)
        except ValueError as exc:
            errors.append(f"地址 {address_index} 解析失败：{exc}")
            continue

        for order_index, order_group in enumerate(address_group["order_groups"], start=1):
            order_no = normalize_text(order_group["external_order_no"])
            if not order_no:
                errors.append(f"地址 {address_index} 的外部单号 {order_index} 未填写。")
            elif order_no in seen_order_numbers:
                errors.append(
                    f"外部平台单号 {order_no} 重复出现："
                    f"{seen_order_numbers[order_no]} 与 地址 {address_index} 的外部单号 {order_index}。"
                )
            else:
                seen_order_numbers[order_no] = f"地址 {address_index} 的外部单号 {order_index}"
            for item_index, item in enumerate(order_group["line_items"], start=1):
                if not normalize_text(item["sku_code"]):
                    errors.append(f"地址 {address_index} / 外部单号 {order_index} / 商品 {item_index} 未匹配到有效产品。")
                if int(item["quantity"]) <= 0:
                    errors.append(f"地址 {address_index} / 外部单号 {order_index} / 商品 {item_index} 数量必须大于 0。")
    return errors


def validate_batch_rows(batch_rows: list[dict[str, str]], batch_shared_line_items: list[dict[str, object]]) -> list[str]:
    errors: list[str] = []
    seen_order_numbers: dict[str, str] = {}
    for row_index, row in enumerate(batch_rows, start=1):
        order_no = normalize_text(row["external_order_no"])
        if not order_no:
            errors.append(f"批量地址第 {row_index} 行缺少外部平台单号。")
        elif order_no in seen_order_numbers:
            errors.append(
                f"批量地址第 {row_index} 行的外部平台单号 {order_no} 重复，已在第 {seen_order_numbers[order_no]} 行出现。"
            )
        else:
            seen_order_numbers[order_no] = str(row_index)
        try:
            split_address(row["address_text"])
        except ValueError as exc:
            errors.append(f"批量地址第 {row_index} 行解析失败：{exc}")

    for item_index, item in enumerate(batch_shared_line_items, start=1):
        if not normalize_text(item["sku_code"]):
            errors.append(f"公共商品 {item_index} 未匹配到有效产品。")
        if int(item["quantity"]) <= 0:
            errors.append(f"公共商品 {item_index} 数量必须大于 0。")
    return errors
