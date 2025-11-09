import os
import sys
import unittest
import filecmp
from excel_extractor import ExcelExtractor

class TestModuleExtractor(unittest.TestCase):
    TEST_DIR = "tests"
    OUTPUT_DIR = os.path.join(TEST_DIR, "results")
    EXPECTED_DIR = os.path.join(TEST_DIR, "expected")

    SPECIFIC_TEST = None  # optional, set from command line

    def setUp(self):
        os.makedirs(self.OUTPUT_DIR, exist_ok=True)

    def test_json_files(self):
        for file_name in os.listdir(self.TEST_DIR):
            if not file_name.endswith(".json"):
                continue
            if self.SPECIFIC_TEST and file_name != self.SPECIFIC_TEST:
                continue

            json_path = os.path.join(self.TEST_DIR, file_name)
            base_name = os.path.splitext(file_name)[0]

            if file_name.endswith("_must_fail.json"):
                # replace .json with .txt instead of appending another _must_fail
                expected_file = os.path.join(self.EXPECTED_DIR, f"{base_name}.txt")
                if not os.path.isfile(expected_file):
                    raise RuntimeError(f"Expected must-fail file missing: {expected_file}")

                generated_file = os.path.join(self.OUTPUT_DIR, f"{base_name}.txt")
                try:
                    extractor = ExcelExtractor(json_path, self.OUTPUT_DIR)
                    extractor.run()
                except Exception as e:
                    with open(generated_file, "w") as f:
                        f.write(str(e))
                else:
                    self.fail(f"{file_name} did not raise an exception")

            else:
                expected_file = os.path.join(self.EXPECTED_DIR, f"{base_name}.csv")
                generated_file = os.path.join(self.OUTPUT_DIR, f"{base_name}.csv")
                extractor = ExcelExtractor(json_path, self.OUTPUT_DIR)
                extractor.run()
                self.assertTrue(os.path.isfile(generated_file), f"{generated_file} not created")
                self.assertTrue(os.path.isfile(expected_file), f"{expected_file} missing")

            # Compare generated output (CSV or must-fail txt) with expected
            self.assertTrue(
                filecmp.cmp(generated_file, expected_file, shallow=False),
                f"{generated_file} does not match expected {expected_file}"
            )


if __name__ == "__main__":
    # optionally pass a test file name: python test_module.py test3.json
    if len(sys.argv) > 1:
        TestModuleExtractor.SPECIFIC_TEST = sys.argv[1]
        test_path = os.path.join(TestModuleExtractor.TEST_DIR, TestModuleExtractor.SPECIFIC_TEST)
        if not os.path.isfile(test_path):
            raise FileNotFoundError(f"Specified test file does not exist: {test_path}")
        # remove argument so unittest doesn't see it
        sys.argv = [sys.argv[0]]
    unittest.main()
