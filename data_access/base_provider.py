# data_access/base_provider.py

from abc import ABC, abstractmethod
from typing import Dict, List, Any

class BaseDataProvider(ABC):
    """
    数据提供者的抽象基类。
    用于从各种源(Excel/CSV/DB...)获取数据,
    并返回统一的 {sheetName: [row_dict, row_dict, ...]} 结构。
    """

    @abstractmethod
    def read_data(self) -> Dict[str, List[Dict[str, Any]]]:
        """
        返回形如:
        {
            "Sheet1": [
                {"[A]": "some value", "[B]": 123, ...},
                ...
            ],
            "Sheet2": [...],
            ...
        }
        """
        pass