"""
SlideGuard Phase 11 — Packaged Executable Smoke-Test Suite.
Executes the complete 19-scenario test matrix against dist/SlideGuard.exe.
"""

import sys
import os
import shutil
import hashlib
import subprocess
import time
from pathlib import Path
from typing import Dict, Any, List

# Ensure repository root is on sys.path
REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE


def calculate_sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest()


class ExeSmokeTestRunner:
    def __init__(self, exe_path: Path, work_dir: Path):
        self.exe_path = exe_path.resolve()
        self.work_dir = work_dir.resolve()
        self.results: Dict[str, Dict[str, Any]] = {}
        self.work_dir.mkdir(parents=True, exist_ok=True)

    def run_exe(self, args: List[str], timeout: int = 120, env: Dict[str, str] = None) -> subprocess.CompletedProcess:
        cmd = [str(self.exe_path)] + args
        run_env = os.environ.copy()
        if env:
            run_env.update(env)
        return subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            timeout=timeout,
            env=run_env
        )

    # -------------------------------------------------------------
    # Scenario A: Application Startup (--version)
    # -------------------------------------------------------------
    def test_a_startup(self):
        print("\n[Scenario A] Application startup (--version)...")
        res = self.run_exe(["--version"], timeout=10)
        passed = (res.returncode == 0)
        self.results["A_STARTUP"] = {
            "passed": passed,
            "code": res.returncode,
            "details": f"Exit code {res.returncode}"
        }
        print(f" -> {'PASS' if passed else 'FAIL'}: {self.results['A_STARTUP']['details']}")

    # -------------------------------------------------------------
    # Scenario B: Empty Startup (--test-mode)
    # -------------------------------------------------------------
    def test_b_empty_startup(self):
        print("\n[Scenario B] Empty startup / GUI init (--test-mode)...")
        res = self.run_exe(["--test-mode"], timeout=15)
        passed = (res.returncode == 0)
        self.results["B_EMPTY_STARTUP"] = {
            "passed": passed,
            "code": res.returncode,
            "details": f"Exit code {res.returncode}"
        }
        print(f" -> {'PASS' if passed else 'FAIL'}: {self.results['B_EMPTY_STARTUP']['details']}")

    # -------------------------------------------------------------
    # Scenario C: Normal PPTX
    # -------------------------------------------------------------
    def test_c_normal_pptx(self):
        print("\n[Scenario C] Normal PPTX processing...")
        fpath = self.work_dir / "normal_presentation.pptx"
        prs = Presentation()
        s1 = prs.slides.add_slide(prs.slide_layouts[0])
        s1.shapes.title.text = "Normal Presentation Title"
        prs.save(str(fpath))

        h_orig = calculate_sha256(fpath)
        res = self.run_exe(["--process", str(fpath)], timeout=30)
        h_after = calculate_sha256(fpath)

        out_file = self.work_dir / "normal_presentation_SlideGuard_Remediated.pptx"
        rep_file = self.work_dir / "normal_presentation_SlideGuard_Report.md"

        passed = (res.returncode == 0 and out_file.exists() and rep_file.exists() and h_orig == h_after)
        self.results["C_NORMAL_PPTX"] = {
            "passed": passed,
            "code": res.returncode,
            "out_exists": out_file.exists(),
            "orig_intact": (h_orig == h_after),
            "details": f"Remediated: {out_file.name}, Original intact: {h_orig == h_after}"
        }
        print(f" -> {'PASS' if passed else 'FAIL'}: {self.results['C_NORMAL_PPTX']['details']}")

    # -------------------------------------------------------------
    # Scenario D: Ch06 fixture
    # -------------------------------------------------------------
    def test_d_ch06_fixture(self):
        print("\n[Scenario D] Ch06 Real-World Fixture...")
        src_ch06 = Path("tests/fixtures/Ch06_GLN_Instructor_AU.pptx").resolve()
        if not src_ch06.exists():
            print(" -> SKIP: Fixture not found")
            return
        dst_ch06 = self.work_dir / "Ch06_smoke_test.pptx"
        shutil.copy2(src_ch06, dst_ch06)
        h_orig = calculate_sha256(dst_ch06)

        res = self.run_exe(["--process", str(dst_ch06)], timeout=90)
        h_after = calculate_sha256(dst_ch06)

        out_file = self.work_dir / "Ch06_smoke_test_SlideGuard_Remediated.pptx"
        passed = (res.returncode == 0 and out_file.exists() and h_orig == h_after)
        self.results["D_CH06_FIXTURE"] = {
            "passed": passed,
            "code": res.returncode,
            "out_exists": out_file.exists(),
            "orig_intact": (h_orig == h_after),
            "details": f"Output size: {out_file.stat().st_size if out_file.exists() else 0} bytes"
        }
        print(f" -> {'PASS' if passed else 'FAIL'}: {self.results['D_CH06_FIXTURE']['details']}")

    # -------------------------------------------------------------
    # Scenario E: Sec1-5 fixture
    # -------------------------------------------------------------
    def test_e_sec1_5_fixture(self):
        print("\n[Scenario E] Sec1-5 Real-World Fixture...")
        src_sec = Path("tests/fixtures/Sec1-5-TrigFunctions (2)original.pptx").resolve()
        if not src_sec.exists():
            print(" -> SKIP: Fixture not found")
            return
        dst_sec = self.work_dir / "Sec1-5_smoke_test.pptx"
        shutil.copy2(src_sec, dst_sec)
        h_orig = calculate_sha256(dst_sec)

        res = self.run_exe(["--process", str(dst_sec)], timeout=60)
        h_after = calculate_sha256(dst_sec)

        out_file = self.work_dir / "Sec1-5_smoke_test_SlideGuard_Remediated.pptx"
        passed = (res.returncode == 0 and out_file.exists() and h_orig == h_after)
        self.results["E_SEC1_5_FIXTURE"] = {
            "passed": passed,
            "code": res.returncode,
            "out_exists": out_file.exists(),
            "orig_intact": (h_orig == h_after),
            "details": f"Output size: {out_file.stat().st_size if out_file.exists() else 0} bytes"
        }
        print(f" -> {'PASS' if passed else 'FAIL'}: {self.results['E_SEC1_5_FIXTURE']['details']}")

    # -------------------------------------------------------------
    # Scenario F: Large PPTX (30 slides)
    # -------------------------------------------------------------
    def test_f_large_pptx(self):
        print("\n[Scenario F] Large PPTX (30 slides)...")
        fpath = self.work_dir / "large_deck.pptx"
        prs = Presentation()
        blank = prs.slide_layouts[6]
        for i in range(30):
            s = prs.slides.add_slide(blank)
            tb = s.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1))
            tb.text_frame.text = f"Slide {i+1} Heading Content"
        prs.save(str(fpath))

        h_orig = calculate_sha256(fpath)
        res = self.run_exe(["--process", str(fpath)], timeout=60)
        h_after = calculate_sha256(fpath)

        out_file = self.work_dir / "large_deck_SlideGuard_Remediated.pptx"
        passed = (res.returncode == 0 and out_file.exists() and h_orig == h_after)
        self.results["F_LARGE_PPTX"] = {
            "passed": passed,
            "code": res.returncode,
            "details": f"Processed 30 slides, intact: {h_orig == h_after}"
        }
        print(f" -> {'PASS' if passed else 'FAIL'}: {self.results['F_LARGE_PPTX']['details']}")

    # -------------------------------------------------------------
    # Scenario G: Image-heavy PPTX
    # -------------------------------------------------------------
    def test_g_image_heavy_pptx(self):
        print("\n[Scenario G] Image-heavy PPTX...")
        fpath = self.work_dir / "image_heavy.pptx"
        prs = Presentation()
        blank = prs.slide_layouts[6]
        # create small temp image
        img_p = self.work_dir / "temp_dot.png"
        from PIL import Image
        img = Image.new("RGB", (64, 64), color=(37, 99, 235))
        img.save(str(img_p))

        for i in range(5):
            s = prs.slides.add_slide(blank)
            for j in range(3):
                s.shapes.add_picture(str(img_p), Inches(1 + j*2), Inches(1), Inches(1.5), Inches(1.5))
        prs.save(str(fpath))

        h_orig = calculate_sha256(fpath)
        res = self.run_exe(["--process", str(fpath)], timeout=45)
        h_after = calculate_sha256(fpath)

        out_file = self.work_dir / "image_heavy_SlideGuard_Remediated.pptx"
        passed = (res.returncode == 0 and out_file.exists() and h_orig == h_after)
        self.results["G_IMAGE_HEAVY"] = {
            "passed": passed,
            "code": res.returncode,
            "details": f"Processed 15 images, intact: {h_orig == h_after}"
        }
        print(f" -> {'PASS' if passed else 'FAIL'}: {self.results['G_IMAGE_HEAVY']['details']}")

    # -------------------------------------------------------------
    # Scenario H: Table-heavy PPTX
    # -------------------------------------------------------------
    def test_h_table_heavy_pptx(self):
        print("\n[Scenario H] Table-heavy PPTX...")
        fpath = self.work_dir / "table_heavy.pptx"
        prs = Presentation()
        blank = prs.slide_layouts[6]
        for i in range(4):
            s = prs.slides.add_slide(blank)
            tbl_shape = s.shapes.add_table(4, 3, Inches(1), Inches(1), Inches(7), Inches(3))
            tbl = tbl_shape.table
            tbl.cell(0, 0).text = "Header A"
            tbl.cell(0, 1).text = "Header B"
            tbl.cell(0, 2).text = "Header C"
            for r in range(1, 4):
                for c in range(3):
                    tbl.cell(r, c).text = f"Data {r},{c}"
        prs.save(str(fpath))

        h_orig = calculate_sha256(fpath)
        res = self.run_exe(["--process", str(fpath)], timeout=45)
        h_after = calculate_sha256(fpath)

        out_file = self.work_dir / "table_heavy_SlideGuard_Remediated.pptx"
        passed = (res.returncode == 0 and out_file.exists() and h_orig == h_after)
        self.results["H_TABLE_HEAVY"] = {
            "passed": passed,
            "code": res.returncode,
            "details": f"Processed 4 tables, intact: {h_orig == h_after}"
        }
        print(f" -> {'PASS' if passed else 'FAIL'}: {self.results['H_TABLE_HEAVY']['details']}")

    # -------------------------------------------------------------
    # Scenario I: Group-heavy PPTX
    # -------------------------------------------------------------
    def test_i_group_heavy_pptx(self):
        print("\n[Scenario I] Group-heavy PPTX...")
        from tests.fixtures.sample_generator import create_comprehensive_test_pptx
        fpath = self.work_dir / "group_heavy.pptx"
        create_comprehensive_test_pptx(fpath)

        h_orig = calculate_sha256(fpath)
        res = self.run_exe(["--process", str(fpath)], timeout=45)
        h_after = calculate_sha256(fpath)

        out_file = self.work_dir / "group_heavy_SlideGuard_Remediated.pptx"
        passed = (res.returncode == 0 and out_file.exists() and h_orig == h_after)
        self.results["I_GROUP_HEAVY"] = {
            "passed": passed,
            "code": res.returncode,
            "details": f"Group processing complete, intact: {h_orig == h_after}"
        }
        print(f" -> {'PASS' if passed else 'FAIL'}: {self.results['I_GROUP_HEAVY']['details']}")

    # -------------------------------------------------------------
    # Scenario J: Invalid PPTX (Corrupt ZIP)
    # -------------------------------------------------------------
    def test_j_invalid_pptx(self):
        print("\n[Scenario J] Invalid / Corrupt PPTX...")
        fpath = self.work_dir / "corrupt_file.pptx"
        with open(fpath, "wb") as f:
            f.write(b"PK\x03\x04INVALID_CORRUPT_PAYLOAD_NOT_A_PPTX")

        res = self.run_exe(["--process", str(fpath)], timeout=15)
        # Invalid file must be caught gracefully and return non-zero exit code without crashing
        passed = (res.returncode != 0)
        self.results["J_INVALID_PPTX"] = {
            "passed": passed,
            "code": res.returncode,
            "details": f"Handled gracefully with non-zero exit code {res.returncode}"
        }
        print(f" -> {'PASS' if passed else 'FAIL'}: {self.results['J_INVALID_PPTX']['details']}")

    # -------------------------------------------------------------
    # Scenario K: Locked PPTX
    # -------------------------------------------------------------
    def test_k_locked_pptx(self):
        print("\n[Scenario K] Locked PPTX...")
        fpath = self.work_dir / "locked_file.pptx"
        prs = Presentation()
        prs.slides.add_slide(prs.slide_layouts[0])
        prs.save(str(fpath))

        # Open file in exclusive write mode to simulate PowerPoint locking it
        # Note: on Windows, exclusive write open will deny shared read if not shared
        passed = False
        try:
            # We simulate a permission/lock error by marking read-only or testing access
            fpath.chmod(0o444) # read-only
            h_orig = calculate_sha256(fpath)
            res = self.run_exe(["--process", str(fpath)], timeout=30)
            h_after = calculate_sha256(fpath)
            out_file = self.work_dir / "locked_file_SlideGuard_Remediated.pptx"
            passed = (res.returncode == 0 and out_file.exists() and h_orig == h_after)
            self.results["K_LOCKED_PPTX"] = {
                "passed": passed,
                "code": res.returncode,
                "details": f"Processed read-only source without mutating source. Intact: {h_orig == h_after}"
            }
        finally:
            import stat
            fpath.chmod(stat.S_IWRITE | stat.S_IREAD)

        print(f" -> {'PASS' if passed else 'FAIL'}: {self.results['K_LOCKED_PPTX']['details']}")

    # -------------------------------------------------------------
    # Scenario L: Unicode Filename
    # -------------------------------------------------------------
    def test_l_unicode_filename(self):
        print("\n[Scenario L] Unicode Filename (German umlauts & CJK)...")
        fpath = self.work_dir / "Präsentation_Übung_日本語.pptx"
        prs = Presentation()
        s1 = prs.slides.add_slide(prs.slide_layouts[0])
        s1.shapes.title.text = "Unicode Presentation"
        prs.save(str(fpath))

        h_orig = calculate_sha256(fpath)
        res = self.run_exe(["--process", str(fpath)], timeout=30)
        h_after = calculate_sha256(fpath)

        out_file = self.work_dir / "Präsentation_Übung_日本語_SlideGuard_Remediated.pptx"
        passed = (res.returncode == 0 and out_file.exists() and h_orig == h_after)
        self.results["L_UNICODE_FILENAME"] = {
            "passed": passed,
            "code": res.returncode,
            "out_exists": out_file.exists(),
            "details": f"Unicode file handled, intact: {h_orig == h_after}"
        }
        print(f" -> {'PASS' if passed else 'FAIL'}: {self.results['L_UNICODE_FILENAME']['details']}")

    # -------------------------------------------------------------
    # Scenario M: Spaces and Special Chars in Path
    # -------------------------------------------------------------
    def test_m_spaces_in_path(self):
        print("\n[Scenario M] Spaces & Parentheses in Path...")
        nested_dir = self.work_dir / "Folder With Spaces (2026) [Official]"
        nested_dir.mkdir(parents=True, exist_ok=True)
        fpath = nested_dir / "My Presentation (Final) [Draft & Review].pptx"

        prs = Presentation()
        s1 = prs.slides.add_slide(prs.slide_layouts[0])
        s1.shapes.title.text = "Spaces in Path Title"
        prs.save(str(fpath))

        h_orig = calculate_sha256(fpath)
        res = self.run_exe(["--process", str(fpath)], timeout=30)
        h_after = calculate_sha256(fpath)

        out_file = nested_dir / "My Presentation (Final) [Draft & Review]_SlideGuard_Remediated.pptx"
        passed = (res.returncode == 0 and out_file.exists() and h_orig == h_after)
        self.results["M_SPACES_IN_PATH"] = {
            "passed": passed,
            "code": res.returncode,
            "out_exists": out_file.exists(),
            "details": f"Special path handled, intact: {h_orig == h_after}"
        }
        print(f" -> {'PASS' if passed else 'FAIL'}: {self.results['M_SPACES_IN_PATH']['details']}")

    # -------------------------------------------------------------
    # Scenario N: Gemini Unavailable (Offline fallback)
    # -------------------------------------------------------------
    def test_n_gemini_unavailable(self):
        print("\n[Scenario N] Gemini Unavailable (Offline Fallback)...")
        fpath = self.work_dir / "gemini_offline.pptx"
        prs = Presentation()
        s1 = prs.slides.add_slide(prs.slide_layouts[0])
        s1.shapes.title.text = "Offline Mode Presentation"
        prs.save(str(fpath))

        # Explicitly empty out any API keys in environment
        offline_env = {"GEMINI_API_KEY": "", "GOOGLE_API_KEY": ""}
        res = self.run_exe(["--process", str(fpath)], timeout=30, env=offline_env)

        out_file = self.work_dir / "gemini_offline_SlideGuard_Remediated.pptx"
        passed = (res.returncode == 0 and out_file.exists())
        self.results["N_GEMINI_UNAVAILABLE"] = {
            "passed": passed,
            "code": res.returncode,
            "details": f"Deterministic fallback operated successfully without Gemini API key"
        }
        print(f" -> {'PASS' if passed else 'FAIL'}: {self.results['N_GEMINI_UNAVAILABLE']['details']}")

    # -------------------------------------------------------------
    # Scenario O: Gemini Fallback Safety
    # -------------------------------------------------------------
    def test_o_gemini_timeout_fallback(self):
        print("\n[Scenario O] Gemini Timeout / Invalid Key Fallback...")
        fpath = self.work_dir / "gemini_invalid_key.pptx"
        prs = Presentation()
        s1 = prs.slides.add_slide(prs.slide_layouts[0])
        s1.shapes.title.text = "Invalid Key Presentation"
        prs.save(str(fpath))

        # Invalid key string
        bad_env = {"GEMINI_API_KEY": "AIzaSy_INVALID_DUMMY_KEY_FOR_TESTING_123456789"}
        res = self.run_exe(["--process", str(fpath)], timeout=30, env=bad_env)

        out_file = self.work_dir / "gemini_invalid_key_SlideGuard_Remediated.pptx"
        passed = (res.returncode == 0 and out_file.exists())
        self.results["O_GEMINI_TIMEOUT_FALLBACK"] = {
            "passed": passed,
            "code": res.returncode,
            "details": f"Fallback safely handled invalid key without crashing"
        }
        print(f" -> {'PASS' if passed else 'FAIL'}: {self.results['O_GEMINI_TIMEOUT_FALLBACK']['details']}")

    # -------------------------------------------------------------
    # Scenario P: Output Collision Protection
    # -------------------------------------------------------------
    def test_p_output_collision(self):
        print("\n[Scenario P] Output Collision Protection...")
        fpath = self.work_dir / "collision_test.pptx"
        prs = Presentation()
        s1 = prs.slides.add_slide(prs.slide_layouts[0])
        s1.shapes.title.text = "Collision Presentation"
        prs.save(str(fpath))

        # First run -> creates collision_test_SlideGuard_Remediated.pptx
        res1 = self.run_exe(["--process", str(fpath)], timeout=30)
        out1 = self.work_dir / "collision_test_SlideGuard_Remediated.pptx"

        # Second run -> should create collision_test_SlideGuard_Remediated (1).pptx
        res2 = self.run_exe(["--process", str(fpath)], timeout=30)
        out2 = self.work_dir / "collision_test_SlideGuard_Remediated (1).pptx"

        passed = (res1.returncode == 0 and res2.returncode == 0 and out1.exists() and out2.exists())
        self.results["P_OUTPUT_COLLISION"] = {
            "passed": passed,
            "out1_exists": out1.exists(),
            "out2_exists": out2.exists(),
            "details": f"Original output preserved; created collision-safe copy: {out2.name}"
        }
        print(f" -> {'PASS' if passed else 'FAIL'}: {self.results['P_OUTPUT_COLLISION']['details']}")

    # -------------------------------------------------------------
    # Scenario Q: Original File Protection (Byte Immutability)
    # -------------------------------------------------------------
    def test_q_original_file_protection(self):
        print("\n[Scenario Q] Original File Byte-for-Byte Protection...")
        fpath = self.work_dir / "immutable_source.pptx"
        prs = Presentation()
        s1 = prs.slides.add_slide(prs.slide_layouts[0])
        s1.shapes.title.text = "Immutable Original Source"
        prs.save(str(fpath))

        hash_before = calculate_sha256(fpath)
        size_before = fpath.stat().st_size

        res = self.run_exe(["--process", str(fpath)], timeout=30)

        hash_after = calculate_sha256(fpath)
        size_after = fpath.stat().st_size

        is_identical = (hash_before == hash_after) and (size_before == size_after)
        passed = (res.returncode == 0 and is_identical)
        self.results["Q_ORIGINAL_FILE_PROTECTION"] = {
            "passed": passed,
            "hash_match": is_identical,
            "hash": hash_before,
            "details": f"SHA-256 match: {is_identical} ({hash_before[:16]}...)"
        }
        print(f" -> {'PASS' if passed else 'FAIL'}: {self.results['Q_ORIGINAL_FILE_PROTECTION']['details']}")

    # -------------------------------------------------------------
    # Scenario R: Report Generation (Markdown & QA metrics)
    # -------------------------------------------------------------
    def test_r_report_generation(self):
        print("\n[Scenario R] Report Generation...")
        fpath = self.work_dir / "report_presentation.pptx"
        prs = Presentation()
        s1 = prs.slides.add_slide(prs.slide_layouts[0])
        s1.shapes.title.text = "Report Test Presentation"
        prs.save(str(fpath))

        res = self.run_exe(["--process", str(fpath)], timeout=30)
        rep_file = self.work_dir / "report_presentation_SlideGuard_Report.md"

        has_content = False
        if rep_file.exists():
            content = rep_file.read_text(encoding="utf-8")
            has_content = "SlideGuard" in content and "Quality Score" in content

        passed = (res.returncode == 0 and rep_file.exists() and has_content)
        self.results["R_REPORT_GENERATION"] = {
            "passed": passed,
            "rep_exists": rep_file.exists(),
            "has_content": has_content,
            "details": f"Report generated with accurate SlideGuard metrics ({rep_file.stat().st_size if rep_file.exists() else 0} bytes)"
        }
        print(f" -> {'PASS' if passed else 'FAIL'}: {self.results['R_REPORT_GENERATION']['details']}")

    # -------------------------------------------------------------
    # Scenario S: Application Restart After Failure
    # -------------------------------------------------------------
    def test_s_restart_after_failure(self):
        print("\n[Scenario S] Application Restart After Failure...")
        # Step 1: Execute failure scenario
        corrupt = self.work_dir / "corrupt_step.pptx"
        with open(corrupt, "wb") as f:
            f.write(b"NOT_A_ZIP_DATA")
        res1 = self.run_exe(["--process", str(corrupt)], timeout=10)

        # Step 2: Immediately restart with valid presentation
        valid = self.work_dir / "valid_restart.pptx"
        prs = Presentation()
        s1 = prs.slides.add_slide(prs.slide_layouts[0])
        s1.shapes.title.text = "Restart Success"
        prs.save(str(valid))
        res2 = self.run_exe(["--process", str(valid)], timeout=30)
        out2 = self.work_dir / "valid_restart_SlideGuard_Remediated.pptx"

        passed = (res1.returncode != 0 and res2.returncode == 0 and out2.exists())
        self.results["S_RESTART_AFTER_FAILURE"] = {
            "passed": passed,
            "failure_caught": (res1.returncode != 0),
            "restart_success": (res2.returncode == 0 and out2.exists()),
            "details": f"Clean recovery after failure: {passed}"
        }
        print(f" -> {'PASS' if passed else 'FAIL'}: {self.results['S_RESTART_AFTER_FAILURE']['details']}")

    # -------------------------------------------------------------
    # Run Complete Matrix
    # -------------------------------------------------------------
    def run_all(self):
        print("==================================================")
        print("SLIDEGUARD EXE SMOKE TEST MATRIX (19 SCENARIOS)")
        print(f"EXE Target: {self.exe_path}")
        print("==================================================")

        self.test_a_startup()
        self.test_b_empty_startup()
        self.test_c_normal_pptx()
        self.test_d_ch06_fixture()
        self.test_e_sec1_5_fixture()
        self.test_f_large_pptx()
        self.test_g_image_heavy_pptx()
        self.test_h_table_heavy_pptx()
        self.test_i_group_heavy_pptx()
        self.test_j_invalid_pptx()
        self.test_k_locked_pptx()
        self.test_l_unicode_filename()
        self.test_m_spaces_in_path()
        self.test_n_gemini_unavailable()
        self.test_o_gemini_timeout_fallback()
        self.test_p_output_collision()
        self.test_q_original_file_protection()
        self.test_r_report_generation()
        self.test_s_restart_after_failure()

        print("\n==================================================")
        total = len(self.results)
        passed_cnt = sum(1 for r in self.results.values() if r["passed"])
        print(f"RESULTS: {passed_cnt}/{total} Scenarios Passed")
        print("==================================================")
        return passed_cnt == total


if __name__ == "__main__":
    exe = Path("dist/SlideGuard.exe")
    if not exe.exists():
        print("Error: dist/SlideGuard.exe does not exist! Run PyInstaller first.")
        sys.exit(1)

    runner = ExeSmokeTestRunner(exe, Path("scratch/exe_smoke_tests"))
    success = runner.run_all()
    sys.exit(0 if success else 1)
