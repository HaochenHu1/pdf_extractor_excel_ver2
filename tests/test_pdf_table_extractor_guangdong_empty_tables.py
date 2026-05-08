import pdf_table_extractor as pte


def test_extract_guangdong_daily_report_with_empty_tables_adds_warning_and_market_rows():
    text = "二、市场交易情况\n今日市场运行平稳。\n三、其他"

    result = pte.extract_guangdong_daily_report(
        pdf_path="dummy.pdf",
        source_file="广东电力现货市场2025年1月运行日报（01.09）.pdf",
        text=text,
        tables=[],
    )

    assert result.report_type == "guangdong_daily"
    assert result.market_rows
    assert any("市场运行平稳" in row["content"] for row in result.market_rows)
    assert "[WARN] 未提取到表格，仅输出文本类结果" in result.diagnostics


def test_extract_guangdong_daily_report_with_empty_tables_has_no_table_rows():
    result = pte.extract_guangdong_daily_report(
        pdf_path="dummy.pdf",
        source_file="广东电力现货市场2025年1月运行日报（01.09）.pdf",
        text="二、市场交易情况\n内容",
        tables=[],
    )

    assert result.t1_volume_rows == []
    assert result.t1_price_rows == []
    assert result.t2_da_rows == []
    assert result.t2_rt_rows == []
