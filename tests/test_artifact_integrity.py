from artifact_integrity import artifact_identity, approval_is_current


def test_approval_rejects_modified_artifacts_and_unexecuted_checks(tmp_path):
    workbook, master = tmp_path / "book.xlsx", tmp_path / "master.xlsx"
    workbook.write_bytes(b"workbook")
    master.write_bytes(b"master")
    contract, spec = {"version": 1}, {"slug": "example"}
    report = {"validation": {"passed": True, "model_quality": {"functional": {"status": "executed", "passed": True}}}, "artifact_identity": artifact_identity(workbook, master, contract, spec)}
    assert approval_is_current(report, workbook, master, contract, spec)
    assert not approval_is_current(report, workbook, master, {"version": 2}, spec)
    assert not approval_is_current(report, workbook, master, contract, {"slug": "changed"})
    workbook.write_bytes(b"modified")
    assert not approval_is_current(report, workbook, master, contract, spec)
    workbook.write_bytes(b"workbook")
    master.write_bytes(b"modified")
    assert not approval_is_current(report, workbook, master, contract, spec)
    master.write_bytes(b"master")
    report["validation"]["model_quality"]["functional"]["status"] = "declared_not_run"
    assert not approval_is_current(report, workbook, master, contract, spec)
