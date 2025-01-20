import unittest

import pandas as pd

from ..libs.read_excel import get_table


class ReadExcelTestCase(unittest.TestCase):
    
    def test_read_excel(self):
        """Data from open excel file is equal test data."""
        df = get_table("\\tests\\data\\test_table.xlsx")
        test_df = pd.DataFrame(
            {
                "1": ["Igor", "Ivan", "George", "Peter", "Valentine"],
                "2": [22, 26, 17, 30, 45],
                "3": ["January", "February", "March", "April", "May"],
                "4": [None, "some file", None, 26, None]
            }
        )
        self.assertEqual(df, test_df)



if __name__ == "__main__":
    unittest.main()