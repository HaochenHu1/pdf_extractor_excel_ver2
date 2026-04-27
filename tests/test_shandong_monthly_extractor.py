import unittest

from shandong_monthly_extractor import extract_shandong_market_disclosure_monthly_report


class DummyTable:
    def __init__(self, title, df):
        self.title = title
        self.df = df


class ShandongExtractorTest(unittest.TestCase):
    def test_extract_required_fields_from_synthetic_text(self):
        text = (
            "一、电网概览 （一）全省全社会用电情况 全社会用电量500亿千瓦时，第一产业用电量10亿千瓦时，同比增长1.2%，"
            "第二产业用电量300亿千瓦时，同比增长2.3%，第三产业用电量120亿千瓦时，同比增长3.4%。"
            "（二）全省发电机组装机及发电总体情况 全省发电装机总容量10000万千瓦，其中水电1000万千瓦、核电2000万千瓦、"
            "火电3000万千瓦、风电1500万千瓦、太阳能发电装机容量2500万千瓦。全省发电量800亿千瓦时，其中火电400亿千瓦时，太阳能发电量60亿千瓦时。"
            "四、交易组织情况 （三）绿电交易组织情况 组织5次省内绿电交易，10家新能源场站、8家售电公司参与，申报电量2.5亿千瓦时，成交电量2.1亿千瓦时，环境溢价20元/兆瓦时。"
            "五、市场结算情况 （一）发电侧交易结算情况 省内发电侧共结算上网电量600亿千瓦时，合约电量500亿千瓦时。2.跨省跨区交易结算情况 跨省跨区交易结算电量80亿千瓦时。"
            "（二）用电侧交易结算情况 1.批发侧结算总体情况 用电批发侧总结算电量550亿千瓦时。"
            "2.零售侧结算情况 8月份，100家零售用户与20家售电公司、3家虚拟电厂线上签订零售合同，结算电量12.32亿千瓦时，"
            "结算均价（仅包含零售套餐中电能量费用和偏差考核费用，下同）350元/兆瓦时。"
        )

        result = extract_shandong_market_disclosure_monthly_report(
            pdf_path="2025年8月山东电力市场信息披露月报.pdf",
            text=text,
            tables=[],
            diagnostics=[],
        )
        value_by_field = {}
        for row in result.info_rows:
            if row["field"] not in value_by_field and row["value"] is not None:
                value_by_field[row["field"]] = row["value"]

        self.assertEqual(value_by_field.get("第一产业用电量"), "10")
        self.assertEqual(value_by_field.get("第一产业同比增长"), "1.2")
        self.assertEqual(value_by_field.get("全省发电装机总容量"), "10000")
        self.assertEqual(value_by_field.get("火电装机容量"), "3000")
        self.assertEqual(value_by_field.get("全省发电量"), "800")
        self.assertEqual(value_by_field.get("太阳能发电量"), "60")
        self.assertEqual(value_by_field.get("省内绿电交易次数"), "5")
        self.assertEqual(value_by_field.get("成交电量"), "2.1")
        self.assertEqual(value_by_field.get("环境溢价"), "20")
        self.assertEqual(value_by_field.get("省内发电侧共结算上网电量"), "600")
        self.assertEqual(value_by_field.get("跨省跨区交易结算电量"), "80")
        self.assertEqual(value_by_field.get("用电批发侧总结算电量"), "550")
        self.assertEqual(value_by_field.get("零售用户数量"), "100")
        self.assertEqual(value_by_field.get("线上签订零售合同结算电量"), "12.32")
        self.assertEqual(value_by_field.get("结算均价"), "350")


if __name__ == "__main__":
    unittest.main()
