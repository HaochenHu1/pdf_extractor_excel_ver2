import pandas as pd

from guangdong_daily_extractor import (
    extract_market_trading_section_text,
    extract_table1_day_ahead_price,
    extract_table1_day_ahead_volume,
    extract_table2_day_ahead_price,
    extract_table2_realtime_price,
    extract_table_operation_date_from_title,
    is_guangdong_daily_report,
    split_price_and_time,
)


def test_filename_detection():
    assert is_guangdong_daily_report("广东电力现货市场2025年1月运行日报（01.09）.pdf")


def test_date_extract_from_title():
    assert extract_table_operation_date_from_title("表1\n运行日2025-1-9日前交易情况") == "2025-01-09"
    assert extract_table_operation_date_from_title("表2 运行日2025年1月10日现货交易情况") == "2025-01-10"


def test_split_price_and_time():
    assert split_price_and_time("321.2 19:00")[1] == "19:00"
    assert split_price_and_time("321.2(0:00)")[1] == "0:00"
    assert split_price_and_time("321.2（04:45）")[1] == "04:45"


def test_market_section_stop_before_next_heading():
    text = "一、xx\n二、市场交易情况\n（一）内容A\n2.内容B\n三、其他\n更多"
    out = extract_market_trading_section_text(text)
    assert "内容A" in out
    assert "三、其他" not in out


def test_table_extracts_and_continuation_rows():
    df = pd.DataFrame([
        ["（三）日前成交电量", ""],
        ["发电侧（含基数及代购电量）", ""],
        ["燃煤", "11.2"],
        ["燃气", "1.2"],
        ["用电侧", ""],
        ["售电公司", "10.0"],
        ["大用户", "2.4"],
        ["（四）日前成交电价", "", "", "", ""],
        ["发电侧", "300 19:00", "200(12:00)", "250", "1.2%"],
        ["燃煤", "301 20:00", "190（0:00）", "248", "0.2%"],
    ])
    v = extract_table1_day_ahead_volume("a.pdf", "表1 运行日2025-1-9日前交易情况", "2025-01-09", "亿kWh", df)
    p1 = extract_table1_day_ahead_price("a.pdf", "表1", "2025-01-09", "元/MWh", df)
    p2 = extract_table2_day_ahead_price("a.pdf", "表2", "2025-01-10", "元/MWh", df)
    p3 = extract_table2_realtime_price("a.pdf", "表2", "2025-01-10", "元/MWh", df)
    assert any(r["category"] == "燃煤" for r in v)
    assert any(r["metric"] == "最高电价" and r["time"] == "19:00" for r in p1)
    assert len(p2) > 0
    assert len(p3) > 0
