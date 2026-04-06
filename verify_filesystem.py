import os
import shutil
import unittest
from pathlib import Path
import winscript.tools.filesystem as fs

class TestFileSystem(unittest.TestCase):
    def setUp(self):
        self.test_dir = Path("test_fs_dir")
        self.test_dir.mkdir(exist_ok=True)
        self.file1 = self.test_dir / "file1.txt"
        self.file2 = self.test_dir / "file2.txt"
        self.subdir = self.test_dir / "subdir"

    def tearDown(self):
        if self.test_dir.exists():
            shutil.rmtree(self.test_dir)

    def test_write_and_read(self):
        res_write = fs.write_file_text(str(self.file1), "hello winscript")
        self.assertTrue(res_write.startswith("Written"))
        
        res_read = fs.read_file_text(str(self.file1))
        self.assertEqual(res_read, "hello winscript")

    def test_list_dir(self):
        fs.write_file_text(str(self.file1), "test")
        self.subdir.mkdir(exist_ok=True)
        
        res = fs.list_dir(str(self.test_dir))
        self.assertIn("[FILE] file1.txt", res)
        self.assertIn("[DIR] subdir", res)

    def test_copy_and_move(self):
        fs.write_file_text(str(self.file1), "test")
        
        res_copy = fs.copy_file(str(self.file1), str(self.file2))
        self.assertTrue(res_copy.startswith("Copied"))
        self.assertTrue(self.file2.exists())
        
        moved_file = self.test_dir / "moved.txt"
        res_move = fs.move_file(str(self.file2), str(moved_file))
        self.assertTrue(res_move.startswith("Moved"))
        self.assertTrue(moved_file.exists())
        self.assertFalse(self.file2.exists())

    def test_delete_file(self):
        fs.write_file_text(str(self.file1), "test")
        res_del = fs.delete_file(str(self.file1))
        self.assertTrue(res_del.startswith("Deleted"))
        self.assertFalse(self.file1.exists())

    def test_delete_dir_error(self):
        self.subdir.mkdir(exist_ok=True)
        res_del = fs.delete_file(str(self.subdir))
        self.assertTrue(res_del.startswith("ERROR:"))
        self.assertIn("refusing to delete", res_del)

    def test_file_exists(self):
        self.subdir.mkdir(exist_ok=True)
        fs.write_file_text(str(self.file1), "test")
        
        res1 = fs.file_exists(str(self.file1))
        self.assertTrue(res1.startswith("EXISTS (file)"))
        
        res2 = fs.file_exists(str(self.subdir))
        self.assertTrue(res2.startswith("EXISTS (directory)"))
        
        res3 = fs.file_exists(str(self.file2))
        self.assertTrue(res3.startswith("NOT FOUND"))

if __name__ == '__main__':
    unittest.main()