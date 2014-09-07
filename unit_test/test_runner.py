# coding=gb2312
__author__ = 'liziqiang'
import unittest
from unittest import defaultTestLoader

import core_test


if __name__ == "__main__":
    runner = unittest.TextTestRunner()
    suite = unittest.TestSuite()
    suite.addTests(defaultTestLoader.loadTestsFromModule(core_test))
    result = runner.run(suite)
