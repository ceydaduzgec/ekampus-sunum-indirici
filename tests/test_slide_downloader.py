"""
Tests for slide_downloader.py, including the --input batch mode.

Slides are served from a local HTTP server, so no internet access is needed.
CairoSVG's system dependencies must be installed (see README).

Run from the repository root:
    python3 -m unittest discover tests
"""
import os, sys, subprocess, threading, http.server, functools, tempfile, unittest
REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, REPO)
import slide_downloader as sd
from pptx import Presentation

def run_cli(*args, cwd):
    return subprocess.run([sys.executable, os.path.join(REPO, "slide_downloader.py"), *args],
                          cwd=cwd, capture_output=True, text=True)

def write(text, d=None, name="in.txt"):
    d = d or tempfile.mkdtemp(); p = os.path.join(d, name)
    with open(p, "w", encoding="utf-8") as f:
        f.write(text)
    return p

class ParseTests(unittest.TestCase):
    def test_valid(self):
        e = sd.parse_input_file(write("Hafta 1 - Giriş: https://example.com/a/svgs/slide1.svg\n\n"
                                      "Hafta 2 - Konu: https://example.com/b/svgs/slide1.svg\n"))
        self.assertEqual([tuple(x) for x in e], [
            (1, "Hafta 1 - Giriş", "https://example.com/a/svgs/slide1.svg"),
            (3, "Hafta 2 - Konu", "https://example.com/b/svgs/slide1.svg")])

    def test_url_colon_not_separator(self):
        self.assertEqual(sd.parse_input_file(write("Ad: http://host:8080/svgs/slide1.svg\n"))[0].url,
                         "http://host:8080/svgs/slide1.svg")

    def test_errors_listed_with_line_numbers(self):
        p = write("Ad https://e.com/s\nAd:https://e.com/s\n: https://e.com/s\nB: ftp://x/s\n"
                  "C: https://e.com/1\nc: https://e.com/2\n")
        with self.assertRaises(sd.InputFileError) as c: sd.parse_input_file(p)
        m = str(c.exception)
        for frag in ("line 1: missing", "line 2: missing", "line 3: presentation name is empty",
                     "line 4: invalid URL", "line 6: duplicate presentation name (same as line 5)"):
            self.assertIn(frag, m)

    def test_empty_missing_encoding(self):
        with self.assertRaises(sd.InputFileError) as c: sd.parse_input_file(write("\n \n"))
        self.assertIn("no presentations", str(c.exception))
        with self.assertRaises(sd.InputFileError) as c: sd.parse_input_file("/nonexistent/x.txt")
        self.assertIn("not found", str(c.exception))
        p = os.path.join(tempfile.mkdtemp(), "b.txt")
        with open(p, "wb") as f:
            f.write(b"\xff\xfe\x00bad")
        with self.assertRaises(sd.InputFileError) as c: sd.parse_input_file(p)
        self.assertIn("UTF-8", str(c.exception))

    def test_safe_filename(self):
        self.assertEqual(sd.safe_filename("2024/25 Hafta:1"), "2024_25 Hafta_1")

class CliTests(unittest.TestCase):
    # Two weekly presentations: week A has 3 slides, week B has 2
    @classmethod
    def setUpClass(cls):
        cls.root = tempfile.mkdtemp()
        for week, n in (("A", 3), ("B", 2)):
            os.makedirs(os.path.join(cls.root, week, "svgs"))
            for i in range(1, n + 1):
                w = (100 if week == "A" else 200) + i
                with open(os.path.join(cls.root, week, "svgs", f"slide{i}.svg"), "w") as f:
                    f.write(f'<svg xmlns="http://www.w3.org/2000/svg" width="{w}" height="50">'
                            f'<rect width="100%" height="100%" fill="blue"/></svg>')
        h = functools.partial(http.server.SimpleHTTPRequestHandler, directory=cls.root)
        h.log_message = lambda *a: None
        cls.httpd = http.server.ThreadingHTTPServer(("127.0.0.1", 0), h)
        cls.base = f"http://127.0.0.1:{cls.httpd.server_address[1]}"
        threading.Thread(target=cls.httpd.serve_forever, daemon=True).start()

    @classmethod
    def tearDownClass(cls): cls.httpd.shutdown(); cls.httpd.server_close()

    def widths(self, path):
        return [s.shapes[0].image.size[0] for s in Presentation(path).slides]

    def test_each_line_is_a_presentation(self):
        d = tempfile.mkdtemp()
        write(f"Hafta 1 - Giriş: {self.base}/A/svgs/slide1.svg\nHafta 2: {self.base}/B/svgs/slide\n", d, "slides.txt")
        r = run_cli("--input", "slides.txt", cwd=d)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        self.assertEqual(self.widths(os.path.join(d, "Hafta 1 - Giriş.pptx")), [101, 102, 103])
        self.assertEqual(self.widths(os.path.join(d, "Hafta 2.pptx")), [201, 202])
        self.assertIn("Created 2 of 2", r.stdout)
        self.assertEqual(sorted(os.listdir(d)), ["Hafta 1 - Giriş.pptx", "Hafta 2.pptx", "slides.txt"])

    def test_options_apply_per_presentation(self):
        d = tempfile.mkdtemp()
        write(f"A: {self.base}/A/svgs/slide1.svg\nB: {self.base}/B/svgs/slide1.svg\n", d, "s.txt")
        r = run_cli("--input", "s.txt", "--start", "2", "--format", "png", "--keep-images",
                    "--output-dir", "imgs", "--presentation", "x.pptx", cwd=d)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        self.assertIn("--presentation is ignored", r.stdout)
        self.assertEqual(self.widths(os.path.join(d, "A.pptx")), [102, 103])
        self.assertEqual(self.widths(os.path.join(d, "B.pptx")), [202])
        self.assertEqual(sorted(os.listdir(os.path.join(d, "imgs", "A"))), ["slide2.png", "slide3.png"])
        self.assertFalse(os.path.exists(os.path.join(d, "x.pptx")))

    def test_failed_line_reported_others_continue(self):
        d = tempfile.mkdtemp()
        write(f"Yok: {self.base}/Z/svgs/slide1.svg\nB: {self.base}/B/svgs/slide1.svg\n", d, "s.txt")
        r = run_cli("--input", "s.txt", "--output-dir", "imgs", cwd=d)
        self.assertEqual(r.returncode, 1)
        self.assertEqual(self.widths(os.path.join(d, "B.pptx")), [201, 202])
        self.assertIn("Created 1 of 2", r.stdout); self.assertIn("line 1: Yok", r.stdout)
        self.assertFalse(os.path.exists(os.path.join(d, "imgs")))

    def test_legacy_usage_unchanged(self):
        d = tempfile.mkdtemp()
        r = run_cli(f"{self.base}/A/svgs/slide1.svg", "--output-dir", "slaytlarim", "--presentation", "dersim.pptx", cwd=d)
        self.assertEqual(r.returncode, 0, r.stderr)
        self.assertEqual(self.widths(os.path.join(d, "dersim.pptx")), [101, 102, 103])
        self.assertEqual(os.listdir(d), ["dersim.pptx"])
        d = tempfile.mkdtemp()
        r = run_cli(f"{self.base}/B/svgs/slide", "--start", "2", "--max", "1", cwd=d)
        self.assertEqual(self.widths(os.path.join(d, "presentation.pptx")), [202])

    def test_arg_errors(self):
        d = tempfile.mkdtemp()
        r = run_cli(cwd=d); self.assertEqual(r.returncode, 2); self.assertIn("either base_url or --input", r.stderr)
        r = run_cli("http://x/slide", "--input", "a.txt", cwd=d)
        self.assertEqual(r.returncode, 2); self.assertIn("not both", r.stderr)
        r = run_cli("--input", "missing.txt", cwd=d)
        self.assertEqual(r.returncode, 1); self.assertIn("Input file not found", r.stderr)
        write("no separator\n", d, "bad.txt")
        r = run_cli("--input", "bad.txt", cwd=d)
        self.assertEqual(r.returncode, 1); self.assertIn("line 1: missing ': ' separator", r.stderr)
        self.assertNotIn("Slide Downloader", r.stdout)

if __name__ == "__main__":
    unittest.main(verbosity=2)
