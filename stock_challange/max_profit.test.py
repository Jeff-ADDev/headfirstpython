import unittest
from max_profit import max_profit

class TestMaxProfit(unittest.TestCase):
    
    def test_max_profit(self):
        self.assertEqual(max_profit([7,1,5,3,6,4]), 5)
        self.assertEqual(max_profit([7,6,4,3,1]), 0)
        self.assertEqual(max_profit([1,2,3,4,5]), 4)
        self.assertEqual(max_profit([2,4,1]), 2)
        self.assertEqual(max_profit([2,1,2,0,1]), 1)

if __name__ == '__main__':
    unittest.main()