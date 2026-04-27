import unittest

from shandong_monthly_extractor import (
    build_shandong_info_dataframe,
    extract_shandong_market_disclosure_monthly_report,
    normalize_shandong_text_for_regex,
    remove_shandong_watermarks,
)


class ShandongExtractorTest(unittest.TestCase):
    def test_extract_with_line_break_watermark_and_footnote_pollution(self):
        text = (
            "晶科慧能 2025年8月12日 10：35：27\n"
            "一、电网概览\n"
            "（一）全省全社会用电情况\n"
            "第一产业用电量\n12.34 亿千瓦时，同比增长\n5.6%\n"
            "城乡居民生活用电\n量 12.34 亿千瓦时，同比增长 5.6%。\n"
            "（二）全省发电机组装机及发电总体情况\n"
            "全省发电装机总容量\n124881.25 万千瓦。\n"
            "其中火电3000万千瓦，太阳能发电 2500万千瓦。全省发电量800亿千瓦时，太阳能\n发电量123.45亿千瓦时，占比10.2%。\n"
            "四、交易组织情况 （三）绿电交易组织情况 组织5次省内绿电交易，成交电量2.1亿千瓦时，环境溢价20元/兆瓦时。\n"
            "五、市场结算情况 （一）发电侧交易结算情况 省内发电侧共结算上网电量600亿千瓦时。2.跨省跨区交易结算情况 跨省跨区交易结算电量80亿千瓦时。\n"
            "（二）用电侧交易结算情况 1.批发侧结算总体情况 用电批发侧总结算电量550亿千瓦时。\n"
            "2.零售侧结算情况 8月份，100家零售用户与20家售电公司、3家虚拟电厂线上签订零售合同，结算电量12.32亿千瓦时，结算均价（仅包含电能量费用）350元/兆瓦时。"
        )

        result = extract_shandong_market_disclosure_monthly_report(
            pdf_path="2025年8月山东电力市场信息披露月报.pdf",
            text=text,
            tables=[],
            diagnostics=[],
        )
        value_by_field = {}
        notes_by_field = {}
        for row in result.info_rows:
            if row["field"] not in value_by_field and row["value"] is not None:
                value_by_field[row["field"]] = row["value"]
            if row["field"] not in notes_by_field and row.get("notes"):
                notes_by_field[row["field"]] = row.get("notes")

        self.assertEqual(value_by_field.get("第一产业用电量"), "12.34")
        self.assertEqual(value_by_field.get("第一产业同比增长"), "5.6")
        self.assertEqual(value_by_field.get("城乡居民生活用电量"), "12.34")
        self.assertEqual(value_by_field.get("城乡居民生活用电量同比增长"), "5.6")
        self.assertEqual(value_by_field.get("全省发电装机总容量"), "24881.25")
        self.assertEqual(value_by_field.get("太阳能发电装机容量"), "2500")
        self.assertEqual(value_by_field.get("太阳能发电量"), "123.45")
        self.assertIn("疑似脚注数字并入数值", notes_by_field.get("全省发电装机总容量", ""))
        self.assertNotIn("月份", {row["field"] for row in result.info_rows})

        # watermark must be removed before diagnostics creation
        joined_diags = "\n".join(result.diagnostics)
        self.assertNotIn("晶科慧能", joined_diags)

    def test_watermark_removal_helper(self):
        src = "晶科慧能2025年08月12日10：35：27 正文 晶科慧能 2025 年 8 月 12 日\n10:35:27"
        cleaned = remove_shandong_watermarks(src)
        self.assertNotIn("晶科慧能", cleaned)

    def test_info_sheet_headers_chinese_and_no_source_text(self):
        info_rows = [
            {
                "报告月份": "2025-08",
                "section": "一、电网概览",
                "subsection": "（一）全省全社会用电情况",
                "field": "第一产业用电量",
                "value": "12.34",
                "unit": "亿千瓦时",
                "notes": "",
                "source_text": "should_not_be_exported",
            }
        ]
        df = build_shandong_info_dataframe(info_rows)
        self.assertEqual(
            list(df.columns),
            ["报告月份", "一级章节", "二级章节", "指标名称", "数值", "单位", "备注"],
        )
        self.assertNotIn("source_text", df.columns)

    def test_normalize_shandong_text_for_regex_joins_split_number_and_unit(self):
        src = "第一产业用电量\n12.34 亿千瓦时 ， 同比增长\n5.6 %"
        cleaned = normalize_shandong_text_for_regex(src)
        self.assertIn("12.34亿千瓦时", cleaned)
        self.assertIn("5.6%", cleaned)


if __name__ == "__main__":
    unittest.main()
