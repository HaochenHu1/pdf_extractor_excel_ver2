import unittest

import pandas as pd
import numpy as np

from shandong_monthly_extractor import (
    build_shandong_info_dataframe,
    extract_shandong_market_disclosure_monthly_report,
    is_table3_continuation_page,
    normalize_shandong_text_for_regex,
    parse_shandong_table_2_cumulative_trade_only,
    parse_shandong_table_3_spot_trade_across_pages,
    parse_shandong_table_8_market_operation_fee_settlement,
    preprocess_shandong_table_image_for_watermark,
    remove_shandong_watermarks,
)


class DummyTable:
    def __init__(self, page, title, df):
        self.page = page
        self.title = title
        self.df = df


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

    def test_table3_split_across_pages_merge(self):
        df1 = pd.DataFrame(
            [
                ["日期", "发电侧日前出清电量", "用电侧日前出清电量", "日前出清均价", "发电侧实时出清电量", "实时出清均价"],
                ["08月01日", "1.1", "1.2", "320", "0.9", "330"],
                ["08月24日", "2.1", "2.2", "321", "1.9", "331"],
            ]
        )
        df2 = pd.DataFrame(
            [
                ["08月25日", "2.3", "2.4", "322", "2.0", "332"],
                ["08月31日", "2.9", "3.0", "323", "2.1", "333"],
                ["合计", "50", "51", "324", "40", "334"],
            ]
        )
        tables = [
            DummyTable(4, "表3：现货交易情况", df1),
            DummyTable(5, "", df2),
        ]
        diags = []
        out = parse_shandong_table_3_spot_trade_across_pages(tables, None, "", diags, report_month="2025-08")
        self.assertIn("08月01日", out["日期"].tolist())
        self.assertIn("08月24日", out["日期"].tolist())
        self.assertIn("08月25日", out["日期"].tolist())
        self.assertIn("08月31日", out["日期"].tolist())
        self.assertIn("合计", out["日期"].tolist())
        self.assertTrue(any("表3续页" in d for d in diags))

    def test_table8_multiline_cells_anchor_by_serial(self):
        df = pd.DataFrame(
            [
                ["序号", "类别", "费用总额", "分摊返还均价", "分摊返还主体"],
                ["4", "新能源场站偏差收益", "10", "0.1", "主体A"],
                ["", "回收", "", "", "主体A补充"],
                ["7", "用户侧日前申报偏差", "11", "0.2", "主体B"],
                ["", "收益回收", "", "", ""],
            ]
        )
        diags = []
        out = parse_shandong_table_8_market_operation_fee_settlement([DummyTable(8, "表8：市场运行费用总体结算情况", df)], None, "", diags)
        row4 = out[out["序号"] == "4"].iloc[0]
        row7 = out[out["序号"] == "7"].iloc[0]
        self.assertIn("新能源场站偏差收益", row4["类别"])
        self.assertIn("回收", row4["类别"])
        self.assertIn("用户侧日前申报偏差", row7["类别"])
        self.assertIn("收益回收", row7["类别"])

    def test_table8_body_rows_1_to_12_present(self):
        rows = [["序号", "类别", "费用总额", "分摊返还均价", "分摊返还主体"]]
        for i in range(1, 13):
            rows.append([str(i), f"类别{i}", str(i * 10), str(i), f"主体{i}"])
        df = pd.DataFrame(rows)
        diags = []
        out = parse_shandong_table_8_market_operation_fee_settlement([DummyTable(8, "表8：市场运行费用总体结算情况", df)], None, "", diags)
        serials = out["序号"].tolist()
        self.assertEqual(serials, [str(i) for i in range(1, 13)])

    def test_table2_upper_half_only(self):
        df = pd.DataFrame(
            [
                ["（一）中长期累计交易情况", "", ""],
                ["交易品种", "累计合约量", "加权平均电价"],
                ["年度双边协商交易", "10", "320"],
                ["月度双边协商交易", "11", "321"],
                ["（二）中长期交易历史净合约情况", "", ""],
                ["日期", "双边协商交易", "合计"],
                ["不应保留", "99", "999"],
            ]
        )
        diags = []
        out = parse_shandong_table_2_cumulative_trade_only([DummyTable(3, "表2：中长期交易情况", df)], None, "", diags)
        self.assertEqual(len(out), 2)
        self.assertNotIn("不应保留", " ".join(out.astype(str).values.flatten().tolist()))
        self.assertTrue(any("边界" in d or "停止" in d for d in diags))

    def test_is_table3_continuation_page_helper(self):
        ctx = {"inside_table3": True, "expected_cols": 6}
        current_text = "08月25日 08月26日 08月27日 合计"
        current_tables = [DummyTable(5, "", pd.DataFrame([["08月25日", "1", "1", "1", "1", "1"]]))]
        self.assertTrue(is_table3_continuation_page(ctx, current_text, current_tables))

    def test_preprocess_shandong_table_image_for_watermark(self):
        img = np.full((50, 50), 255, dtype=np.uint8)
        img[10:12, 5:45] = 120  # dark text line
        img[20:22, 5:45] = 190  # light watermark line
        cleaned = preprocess_shandong_table_image_for_watermark(img)
        self.assertLess(cleaned[10:12, 5:45].mean(), 160)  # dark text preserved
        self.assertGreater(cleaned[20:22, 5:45].mean(), 240)  # watermark whitened

    def test_watermark_tokens_removed_from_table_cells(self):
        df = pd.DataFrame(
            [
                ["交易品种", "累计合约量", "加权平均电价"],
                ["年度双边协商交易 晶科慧能", "10", "320 2025年9月24日16:36:20"],
            ]
        )
        diags = []
        out = parse_shandong_table_2_cumulative_trade_only([DummyTable(3, "表2：中长期交易情况", df)], None, "", diags)
        flat = " ".join(out.astype(str).values.flatten().tolist())
        self.assertNotIn("晶科慧能", flat)
        self.assertNotIn("16:36:20", flat)


if __name__ == "__main__":
    unittest.main()
