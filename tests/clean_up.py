
"""
Clean up any temporary files created by test.py (sometimes temp files
are inadvertently left if a test fails before cleaning up within test.py
itself).
"""

import test

if __name__ == '__main__':
    fh = test.FileHandler()
    fh.clean_up()
