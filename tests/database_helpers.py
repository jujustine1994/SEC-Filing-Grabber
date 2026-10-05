from database import create_database, read_marker


def make_database(tmp_path, *, test_mode=True):
    root = tmp_path / 'database'
    if not root.exists():
        create_database(root, test_mode=test_mode)
    else:
        assert read_marker(root)['test_mode'] == test_mode
    return root
