"""
JSONエンコーダーのテスト

カスタムJSONエンコーダーが様々なオブジェクトを
適切にシリアライズできることを確認します。
"""
import unittest
import json
import datetime
import numpy as np
import pandas as pd

from xlwings_rpc.utils.json_encoder import XlwingsRpcJSONEncoder, json_dumps


class TestJsonEncoder(unittest.TestCase):
    """カスタムJSONエンコーダーのテスト"""
    
    def test_primitive_types(self):
        """基本的な型のシリアライズをテスト"""
        test_data = {
            "int": 42,
            "float": 3.14,
            "str": "hello",
            "bool": True,
            "none": None,
            "list": [1, 2, 3],
            "dict": {"a": 1, "b": 2}
        }
        
        # カスタムエンコーダーでシリアライズ
        result = json_dumps(test_data)
        
        # 結果を検証
        parsed = json.loads(result)
        self.assertEqual(parsed["int"], 42)
        self.assertEqual(parsed["float"], 3.14)
        self.assertEqual(parsed["str"], "hello")
        self.assertEqual(parsed["bool"], True)  # JSONでは小文字のtrueになるが、パース後はPythonのTrueになる
        self.assertIsNone(parsed["none"])
        self.assertEqual(parsed["list"], [1, 2, 3])
        self.assertEqual(parsed["dict"], {"a": 1, "b": 2})
    
    def test_datetime(self):
        """日付型のシリアライズをテスト"""
        test_data = {
            "datetime": datetime.datetime(2023, 1, 1, 12, 30, 45),
            "date": datetime.date(2023, 1, 1)
        }
        
        # カスタムエンコーダーでシリアライズ
        result = json_dumps(test_data)
        
        # 結果を検証
        parsed = json.loads(result)
        self.assertEqual(parsed["datetime"], "2023-01-01T12:30:45")
        self.assertEqual(parsed["date"], "2023-01-01")
    
    def test_numpy_types(self):
        """NumPy型のシリアライズをテスト"""
        test_data = {
            "int32": np.int32(42),
            "int64": np.int64(42),
            "float32": np.float32(3.14),
            "float64": np.float64(3.14),
            "bool": np.bool_(True),
            "array": np.array([1, 2, 3])
        }
        
        # カスタムエンコーダーでシリアライズ
        result = json_dumps(test_data)
        
        # 結果を検証
        parsed = json.loads(result)
        self.assertEqual(parsed["int32"], 42)
        self.assertEqual(parsed["int64"], 42)
        self.assertAlmostEqual(parsed["float32"], 3.14, places=5)
        self.assertAlmostEqual(parsed["float64"], 3.14)
        self.assertEqual(parsed["bool"], True)
        self.assertEqual(parsed["array"], [1, 2, 3])
    
    def test_pandas_dataframe(self):
        """pandas DataFrameのシリアライズをテスト"""
        df = pd.DataFrame({
            'A': [1, 2, 3],
            'B': ['a', 'b', 'c']
        })
        
        test_data = {
            "dataframe": df
        }
        
        # カスタムエンコーダーでシリアライズ
        result = json_dumps(test_data)
        
        # 結果を検証
        parsed = json.loads(result)
        self.assertEqual(parsed["dataframe"]["type"], "dataframe")
        self.assertEqual(parsed["dataframe"]["columns"], ["A", "B"])
        self.assertEqual(len(parsed["dataframe"]["data"]), 3)
        self.assertEqual(parsed["dataframe"]["data"][0][0], 1)
        self.assertEqual(parsed["dataframe"]["data"][0][1], "a")
    
    def test_complex_nested_structure(self):
        """複雑なネストされた構造のシリアライズをテスト"""
        test_data = {
            "level1": {
                "level2": {
                    "array": np.array([1, 2, 3]),
                    "date": datetime.date(2023, 1, 1)
                },
                "list": [
                    np.int64(42),
                    {
                        "nested": np.float32(3.14)
                    }
                ]
            }
        }
        
        # カスタムエンコーダーでシリアライズ
        result = json_dumps(test_data)
        
        # 結果を検証
        parsed = json.loads(result)
        self.assertEqual(parsed["level1"]["level2"]["array"], [1, 2, 3])
        self.assertEqual(parsed["level1"]["level2"]["date"], "2023-01-01")
        self.assertEqual(parsed["level1"]["list"][0], 42)
        self.assertAlmostEqual(parsed["level1"]["list"][1]["nested"], 3.14, places=5)
    
    def test_custom_class_fallback(self):
        """カスタムクラスのフォールバック処理をテスト"""
        # 単純なカスタムクラス
        class TestClass:
            def __init__(self, value):
                self.value = value
            
            def __str__(self):
                return f"TestClass({self.value})"
        
        test_data = {
            "custom": TestClass(42)
        }
        
        # カスタムエンコーダーでシリアライズ
        result = json_dumps(test_data)
        
        # 結果を検証
        parsed = json.loads(result)
        self.assertEqual(parsed["custom"], "TestClass(42)")
    
    def test_version_number_simulation(self):
        """VersionNumberクラスのシミュレーションテスト"""
        # VersionNumberクラスをシミュレート
        class VersionNumber:
            def __init__(self, version):
                self.version = version
            
            def __str__(self):
                return self.version
        
        test_data = {
            "version": VersionNumber("16.84")
        }
        
        # カスタムエンコーダーでシリアライズ
        result = json_dumps(test_data)
        
        # 結果を検証
        parsed = json.loads(result)
        self.assertEqual(parsed["version"], "16.84")


if __name__ == "__main__":
    unittest.main()
