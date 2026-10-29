"""Run with:  python test_clean_sheets.py   (or: pytest)"""
import pandas as pd
import clean_sheets as cs


def test_snake():
    assert cs.snake(" Customer Name ") == "customer_name"
    assert cs.snake("Amount (UGX)") == "amount_ugx"


def test_phone():
    assert cs.clean_phone("0772 123 456") == "+256772123456"
    assert cs.clean_phone("+256 701 234 567") == "+256701234567"
    assert cs.clean_phone("256704111222") == "+256704111222"
    assert cs.clean_phone("12345") is None
    assert cs.clean_phone(None) is None


def test_amount():
    assert cs.clean_amount("UGX 1,200,000") == 1200000.0
    assert cs.clean_amount("85,000") == 85000.0
    assert cs.clean_amount("abc") is None


def test_date():
    assert cs.clean_date("03/09/2026") == pd.Timestamp("2026-09-03")   # day first
    assert cs.clean_date("2026-09-05") == pd.Timestamp("2026-09-05")   # ISO year first
    assert pd.isna(cs.clean_date("not sure"))


def test_full_run(tmp_path):
    out = tmp_path / "report.xlsx"
    merged, summary, issues = cs.run("sample_data", out)
    assert out.exists()
    assert len(merged) == 7                      # 9 raw rows - 1 empty row - 1 duplicate
    assert merged["phone"].dropna().str.startswith("+256").all()
    assert set(pd.ExcelFile(out).sheet_names) == {"Cleaned Data", "Summary", "Issues Log"}
    assert summary["total"].sum() == merged["amount_ugx"].sum()


if __name__ == "__main__":
    import tempfile, pathlib
    test_snake(); test_phone(); test_amount(); test_date()
    test_full_run(pathlib.Path(tempfile.mkdtemp()))
    print("All tests passed")
