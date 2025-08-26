using System;
using System.IO;
using System.Linq;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Windows.Forms;
using System.Collections.Generic;

using OpenCvSharp;
using OpenCvSharp.Extensions;

// 충돌 방지용 별칭
using CvPoint = OpenCvSharp.Point;
using CvSize = OpenCvSharp.Size;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;
using System.Drawing.Drawing2D;              // Paint용
using SDPoint = System.Drawing.Point;
using SDPointF = System.Drawing.PointF;
using SDRectF = System.Drawing.RectangleF;

namespace PCBMissing
{
    public class PCBMissingControl : UserControl
    {
        // ===== 기본 경로/옵션 =====
        private string GOLDEN_PATH = @"pcb4\Data\Images\Normal\0000.jpg";
        private string TEST_PATH = @"pcb4\Data\Images\Anomaly\073.jpg";

        private bool USE_LATEST_IN_TEST_FOLDER = false;
        private string TEST_FOLDER = @"pcb4\Data\Images\Anomaly";
        private readonly string[] TEST_EXTS = new[] { ".jpg", ".jpeg", ".png", ".bmp" };

        // 전역(디폴트) 컷값
        private double SCORE_TH = 0.70;
        private double EDGE_RATIO_TH = 0.65;
        private double BRIGHT_RATIO_TH = 0.70;

        // 전역(디폴트) 검색 옵션
        private bool USE_LEFT_SEARCH = true;
        private double LEFT_WIDTH_RATIO = 0.60;

        // 배율 변동 대응(옵션)
        private bool USE_MULTI_SCALE = false;
        private double SCALE_MIN = 0.90, SCALE_MAX = 1.10, SCALE_STEP = 0.05;

        // 매칭 전처리
        private bool USE_EDGE_MATCH = true;     // 소벨 magnitude 기반 매칭 (평평하면 자동 폴백)

        // 다중 ROI 전체 판정 방식
        private bool ALL_ROIS_MUST_PRESENT = true; // AND(true)/OR(false)

        // ===== Prefs (자동 저장/복원) =====
        private static readonly string PREF_DIR =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "PCBMissing");
        private static readonly string PREF_FILE = Path.Combine(PREF_DIR, "pcb_missing_prefs.txt");

        // ===== UI =====
        private TextBox tbGolden, tbTest, tbFolder;
        private CheckBox cbLatest, cbLeftSearch, cbMultiScale, cbEdgeMatch, cbAllPresent;
        private Button btnPick, btnCheck, btnBrowseG, btnBrowseT, btnBrowseF, btnEditRoi;
        private NumericUpDown numScore, numEdgeR, numBrightR, numLeftRatio;
        private PictureBox pbView;
        private TextBox tbLog;

        // ===== ROI 내장 편집기 =====
        private List<ROIConf> _roiEditing = new List<ROIConf>();

        // 선택 상태
        private List<int> _selSet = new List<int>(); // 여러 개 선택
        private int _lastSel = -1;                   // 마지막 클릭한 하나(핸들/라벨 표시용)
        private bool _mouseDown = false;
        private enum HitType { None, Body, N, S, E, W, NE, NW, SE, SW }
        private HitType _hit = HitType.None;
        private Rect _dragOrigRect;                  // 단일 드래그용
        private Dictionary<int, Rect> _dragOrigRects; // 다중 이동용
        private OpenCvSharp.Point _dragStartImg;

        // 복사/붙여넣기 클립보드
        private List<Rect> _clipboard = new List<Rect>();

        // 뷰(줌/팬)
        private float _zoom = 1f;                    // 배율 (1.0 = 100%)
        private PointF _pan = new PointF(0, 0);      // 이미지 좌표계 기준 뷰 원점
        private bool _panning = false;
        private Point _panStartView;
        private PointF _panStartImg;

        // 골든 이미지(편집용) - 편집 모드에서 PictureBox에 직접 그림
        private Bitmap _goldenBmp;
        private Mat _goldenMatColorForSave;

        // 핸들 크기(px)
        private const int HANDLE = 8;

        private bool _roiEditMode = false;
        // 예전 코드가 roiEditMode(언더스코어 없음)를 참조해도 빌드되도록 브릿지 프로퍼티
        private bool roiEditMode { get => _roiEditMode; set => _roiEditMode = value; }

        // ROI 이동 방지: 골든 ROI 주변만 검색
        private bool LOCK_TO_GOLDEN = true;   // true면 ROI 주변에서만 매칭
        private double LOCK_MARGIN = 0.30;   // ROI의 가로/세로를 기준으로 ±30% 만큼 확장한 영역에서만 검색

        private const bool ROI_EDIT_STARTS_EMPTY = true;

        public PCBMissingControl()
        {
            InitializeUi();
            ApplyDefaultsToUi();
            ApplyTheme();
            LoadPrefs();
        }

        // ---------------- 데이터 구조 (ROI별 설정) ----------------
        private class ROIConf
        {
            public int Index;
            public Rect Roi;
            public string Name;
            // 개별 컷값(널이면 전역값 사용)
            public double? ScoreTh;
            public double? EdgeTh;
            public double? BrightTh;
            // 개별 검색옵션
            public bool? UseLeftSearch;
            public double? LeftRatio;
            public Rect? SearchRect; // 지정 시 최우선
        }

        // ---------------- UI ----------------

        private void InitializeUi()
        {
            this.BackColor = Color.FromArgb(32, 32, 32);

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                BackColor = Color.FromArgb(28, 28, 28)
            };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            this.Controls.Add(root);

            // --- 경로 패널 ---
            var pPath = new TableLayoutPanel { Dock = DockStyle.Top, ColumnCount = 7, AutoSize = true };
            pPath.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            pPath.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            pPath.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            pPath.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            pPath.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            pPath.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            pPath.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            var lbG = MakeLabel("Golden:"); tbGolden = MakeTextBox();
            btnBrowseG = MakeButton("..."); btnBrowseG.Click += (s, e) => { BrowseFile(tbGolden); };

            var lbT = MakeLabel("Test:"); tbTest = MakeTextBox();
            btnBrowseT = MakeButton("..."); btnBrowseT.Click += (s, e) => { BrowseFile(tbTest); };

            pPath.Controls.Add(lbG, 0, 0);
            pPath.Controls.Add(tbGolden, 1, 0);
            pPath.Controls.Add(btnBrowseG, 2, 0);
            pPath.Controls.Add(lbT, 3, 0);
            pPath.Controls.Add(tbTest, 4, 0);
            pPath.Controls.Add(btnBrowseT, 5, 0);

            cbLatest = new CheckBox { Text = "Use Latest in Folder", ForeColor = Color.Gainsboro, AutoSize = true, Margin = new Padding(8, 6, 8, 6) };
            pPath.Controls.Add(cbLatest, 6, 0);

            var lbF = MakeLabel("Folder:"); tbFolder = MakeTextBox(); tbFolder.Width = 280;
            btnBrowseF = MakeButton("..."); btnBrowseF.Click += (s, e) => { BrowseFolder(tbFolder); };

            pPath.Controls.Add(lbF, 0, 1);
            pPath.Controls.Add(tbFolder, 1, 1);
            pPath.Controls.Add(btnBrowseF, 2, 1);

            root.Controls.Add(pPath);

            // --- 옵션/버튼 패널 ---
            var pOpts = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(8), BackColor = Color.FromArgb(26, 26, 26) };

            numScore = MakeNumeric(0.70m, 0, 1, 0.01m); pOpts.Controls.Add(MakeLabeled("Score TH", numScore));
            numEdgeR = MakeNumeric(0.65m, 0, 5, 0.01m); pOpts.Controls.Add(MakeLabeled("Edge Ratio TH", numEdgeR));
            numBrightR = MakeNumeric(0.70m, 0, 5, 0.01m); pOpts.Controls.Add(MakeLabeled("Bright Ratio TH", numBrightR));

            cbLeftSearch = new CheckBox { Text = "Left Search", Checked = true, ForeColor = Color.Gainsboro, AutoSize = true, Margin = new Padding(16, 6, 8, 6) };
            numLeftRatio = MakeNumeric(0.60m, 0.10m, 1.00m, 0.05m);
            pOpts.Controls.Add(cbLeftSearch);
            pOpts.Controls.Add(MakeLabeled("Left Ratio", numLeftRatio));

            cbMultiScale = new CheckBox { Text = "Multi-Scale", Checked = false, ForeColor = Color.Gainsboro, AutoSize = true, Margin = new Padding(16, 6, 8, 6) };
            pOpts.Controls.Add(cbMultiScale);

            cbEdgeMatch = new CheckBox { Text = "Edge Match", Checked = true, ForeColor = Color.Gainsboro, AutoSize = true, Margin = new Padding(16, 6, 8, 6) };
            pOpts.Controls.Add(cbEdgeMatch);

            cbAllPresent = new CheckBox { Text = "All ROIs must be present", Checked = true, ForeColor = Color.Gainsboro, AutoSize = true, Margin = new Padding(16, 6, 8, 6) };
            pOpts.Controls.Add(cbAllPresent);

            btnPick = MakeButton("ROI 편집");
            btnPick.Click += (s, e) => ToggleRoiEdit();

            btnEditRoi = MakeButton("ROI 설정");
            btnEditRoi.Click += (s, e) => EditRoiUi();

            btnCheck = MakeButton("② 검사 실행");
            btnCheck.Click += (s, e) => RunCheckUi();

            pOpts.Controls.Add(btnPick);
            pOpts.Controls.Add(btnEditRoi);
            pOpts.Controls.Add(btnCheck);

            root.Controls.Add(pOpts);

            // --- 뷰/로그 ---
            var pView = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Horizontal, SplitterDistance = 420 };
            pbView = new PictureBox { Dock = DockStyle.Fill, BackColor = Color.Black, SizeMode = PictureBoxSizeMode.Zoom };
            pView.Panel1.Controls.Add(pbView);

            tbLog = new TextBox { Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Vertical, ReadOnly = true, BackColor = Color.FromArgb(20, 20, 20), ForeColor = Color.Gainsboro };
            pView.Panel2.Controls.Add(tbLog);

            root.Controls.Add(pView);

            pbView.TabStop = true;                   // 키 입력 받기
            pbView.Paint += PbView_Paint;
            pbView.MouseDown += PbView_MouseDown;
            pbView.MouseMove += PbView_MouseMove;
            pbView.MouseUp += PbView_MouseUp;
            pbView.MouseWheel += PbView_MouseWheel;
            pbView.MouseEnter += (s, e) => pbView.Focus();
            pbView.PreviewKeyDown += (s, e) => e.IsInputKey = true;
            pbView.KeyDown += PbView_KeyDown;

        }

        private Label MakeLabel(string text) =>
            new Label { Text = text, AutoSize = true, ForeColor = Color.WhiteSmoke, Margin = new Padding(8, 8, 4, 4) };

        private TextBox MakeTextBox() =>
            new TextBox { Width = 320, Margin = new Padding(4, 4, 4, 4), BackColor = Color.FromArgb(30, 30, 30), ForeColor = Color.WhiteSmoke, BorderStyle = BorderStyle.FixedSingle };

        private Button MakeButton(string text)
        {
            var btn = new Button { Text = text, AutoSize = true, Margin = new Padding(8, 4, 4, 4) };
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderColor = Color.FromArgb(90, 90, 90);
            btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(70, 110, 210);
            btn.FlatAppearance.MouseDownBackColor = Color.FromArgb(55, 90, 175);
            btn.UseVisualStyleBackColor = false;
            btn.BackColor = Color.FromArgb(60, 120, 220);
            btn.ForeColor = Color.White;
            btn.Padding = new Padding(8, 4, 8, 4);
            try { btn.Font = new Font(btn.Font, FontStyle.Bold); } catch { }
            return btn;
        }

        private NumericUpDown MakeNumeric(decimal val, decimal min, decimal max, decimal step) =>
            new NumericUpDown { DecimalPlaces = 2, Increment = step, Minimum = min, Maximum = max, Value = val, Width = 80, Margin = new Padding(6), BackColor = Color.FromArgb(30, 30, 30), ForeColor = Color.WhiteSmoke };

        private Control MakeLabeled(string label, Control ctrl)
        {
            var p = new FlowLayoutPanel { AutoSize = true };
            p.Controls.Add(MakeLabel(label));
            p.Controls.Add(ctrl);
            return p;
        }

        private void ApplyDefaultsToUi()
        {
            tbGolden.Text = GOLDEN_PATH;
            tbTest.Text = TEST_PATH;
            tbFolder.Text = TEST_FOLDER;
            cbLatest.Checked = USE_LATEST_IN_TEST_FOLDER;

            numScore.Value = (decimal)SCORE_TH;
            numEdgeR.Value = (decimal)EDGE_RATIO_TH;
            numBrightR.Value = (decimal)BRIGHT_RATIO_TH;
            cbLeftSearch.Checked = USE_LEFT_SEARCH;
            numLeftRatio.Value = (decimal)LEFT_WIDTH_RATIO;
            cbMultiScale.Checked = USE_MULTI_SCALE;
            cbEdgeMatch.Checked = USE_EDGE_MATCH;
            cbAllPresent.Checked = ALL_ROIS_MUST_PRESENT;
        }

        private void ReadUiToOptions()
        {
            GOLDEN_PATH = tbGolden.Text?.Trim();
            TEST_PATH = tbTest.Text?.Trim();
            TEST_FOLDER = tbFolder.Text?.Trim();
            USE_LATEST_IN_TEST_FOLDER = cbLatest.Checked;

            SCORE_TH = (double)numScore.Value;
            EDGE_RATIO_TH = (double)numEdgeR.Value;
            BRIGHT_RATIO_TH = (double)numBrightR.Value;

            USE_LEFT_SEARCH = cbLeftSearch.Checked;
            LEFT_WIDTH_RATIO = (double)numLeftRatio.Value;
            USE_MULTI_SCALE = cbMultiScale.Checked;
            USE_EDGE_MATCH = cbEdgeMatch.Checked;
            ALL_ROIS_MUST_PRESENT = cbAllPresent.Checked;
        }

        private void ApplyTheme()
        {
            this.BackColor = Color.FromArgb(18, 18, 18);
        }

        // ---------------- Prefs ----------------

        private void LoadPrefs()
        {
            try
            {
                Directory.CreateDirectory(PREF_DIR);
                if (!File.Exists(PREF_FILE)) return;

                var dict = File.ReadAllLines(PREF_FILE)
                    .Where(l => !string.IsNullOrWhiteSpace(l) && l.Contains("="))
                    .Select(l => new { k = l.Substring(0, l.IndexOf('=')), v = l.Substring(l.IndexOf('=') + 1) })
                    .GroupBy(x => x.k).ToDictionary(g => g.Key.Trim(), g => g.Last().v.Trim());

                string S(string k, string d) => dict.TryGetValue(k, out var v) && !string.IsNullOrWhiteSpace(v) ? v : d;
                bool B(string k, bool d) => dict.TryGetValue(k, out var v) && bool.TryParse(v, out var b) ? b : d;
                decimal D(string k, decimal d, decimal min, decimal max)
                {
                    if (dict.TryGetValue(k, out var v) &&
                        decimal.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out var val))
                        return Math.Min(max, Math.Max(min, val));
                    return d;
                }

                tbGolden.Text = S("Golden", tbGolden.Text);
                tbTest.Text = S("Test", tbTest.Text);
                tbFolder.Text = S("Folder", tbFolder.Text);

                cbLatest.Checked = B("UseLatest", cbLatest.Checked);
                cbLeftSearch.Checked = B("LeftSearch", cbLeftSearch.Checked);
                cbMultiScale.Checked = B("MultiScale", cbMultiScale.Checked);
                cbEdgeMatch.Checked = B("EdgeMatch", cbEdgeMatch.Checked);
                cbAllPresent.Checked = B("AllPresent", cbAllPresent.Checked);

                numScore.Value = D("ScoreTH", numScore.Value, numScore.Minimum, numScore.Maximum);
                numEdgeR.Value = D("EdgeTH", numEdgeR.Value, numEdgeR.Minimum, numEdgeR.Maximum);
                numBrightR.Value = D("BrightTH", numBrightR.Value, numBrightR.Minimum, numBrightR.Maximum);
                numLeftRatio.Value = D("LeftRatio", numLeftRatio.Value, numLeftRatio.Minimum, numLeftRatio.Maximum);
            }
            catch { }
        }

        private void SavePrefs()
        {
            try
            {
                Directory.CreateDirectory(PREF_DIR);
                var lines = new[]
                {
                    $"Golden={tbGolden.Text}",
                    $"Test={tbTest.Text}",
                    $"Folder={tbFolder.Text}",
                    $"UseLatest={cbLatest.Checked}",
                    $"LeftSearch={cbLeftSearch.Checked}",
                    $"MultiScale={cbMultiScale.Checked}",
                    $"EdgeMatch={cbEdgeMatch.Checked}",
                    $"AllPresent={cbAllPresent.Checked}",
                    $"ScoreTH={numScore.Value.ToString(CultureInfo.InvariantCulture)}",
                    $"EdgeTH={numEdgeR.Value.ToString(CultureInfo.InvariantCulture)}",
                    $"BrightTH={numBrightR.Value.ToString(CultureInfo.InvariantCulture)}",
                    $"LeftRatio={numLeftRatio.Value.ToString(CultureInfo.InvariantCulture)}",
                };
                File.WriteAllLines(PREF_FILE, lines);
            }
            catch { }
        }

        // ---------------- 유틸/로그 ----------------

        private void Log(string msg) => tbLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {msg}\r\n");

        private void BrowseFile(TextBox tb)
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp|All Files|*.*";
                if (File.Exists(tb.Text)) dlg.InitialDirectory = Path.GetDirectoryName(tb.Text);
                if (dlg.ShowDialog() == DialogResult.OK) { tb.Text = dlg.FileName; SavePrefs(); }
            }
        }

        private void BrowseFolder(TextBox tb)
        {
            using (var dlg = new FolderBrowserDialog())
            {
                if (Directory.Exists(tb.Text)) dlg.SelectedPath = tb.Text;
                if (dlg.ShowDialog() == DialogResult.OK) { tb.Text = dlg.SelectedPath; SavePrefs(); }
            }
        }

        private void ShowMatOnPictureBox(Mat mat)
        {
            Bitmap bmp = null;
            try
            {
                bmp = BitmapConverter.ToBitmap(mat);
                var old = pbView.Image;
                pbView.Image = bmp;
                if (old != null) old.Dispose();
            }
            catch
            {
                if (bmp != null) bmp.Dispose();
                throw;
            }
        }

        // ---------------- 버튼 핸들러 ----------------
        private void EditRoiUi()
        {
            try
            {
                var list = LoadRoiListFromFile();
                if (list.Count == 0) { MessageBox.Show("먼저 ROI를 저장해 주세요(① ROI 픽).", "ROI 설정", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }

                using (var dlg = new ROIEditorForm(list))
                {
                    if (dlg.ShowDialog(this.FindForm()) == DialogResult.OK)
                    {
                        SaveRoiListToFile(dlg.Result);
                        Log("ROI 설정 저장 완료 (roi.txt 업데이트)");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ROI 설정 - Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Log("ERROR: " + ex);
            }
        }

        private void RunCheckUi()
        {
            try
            {
                ReadUiToOptions();

                string testPath = TEST_PATH;
                if (USE_LATEST_IN_TEST_FOLDER && Directory.Exists(TEST_FOLDER))
                {
                    var latest = Directory.EnumerateFiles(TEST_FOLDER)
                                          .Where(p => TEST_EXTS.Contains(Path.GetExtension(p).ToLower()))
                                          .Select(p => new FileInfo(p))
                                          .OrderByDescending(f => f.LastWriteTime)
                                          .FirstOrDefault();
                    if (latest != null) testPath = latest.FullName;
                }

                if (!File.Exists(GOLDEN_PATH) || !File.Exists(testPath))
                {
                    MessageBox.Show("골든/테스트 이미지 경로를 확인하세요.", "검사 실행", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                SavePrefs();
                RunCheckCore(GOLDEN_PATH, testPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "검사 실행 - Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Log("ERROR: " + ex);
            }
        }

        // ---------------- ROI 도구 ----------------

        private List<Rect> PickRois(Mat img, string winTitle)
        {
            var rois = new List<Rect>();
            bool dragging = false; CvPoint p0 = new CvPoint();
            Rect cur = new Rect();

            using (var win = new Window(winTitle, WindowFlags.Normal))
            {
                Action draw = () =>
                {
                    using (var tmp = img.Clone())
                    {
                        foreach (var r in rois) Cv2.Rectangle(tmp, r, new Scalar(0, 255, 0), 2);
                        if (dragging && cur.Width > 0 && cur.Height > 0)
                            Cv2.Rectangle(tmp, cur, new Scalar(0, 180, 255), 2);
                        Cv2.PutText(tmp, "Drag to add ROI | R: undo last | ENTER: save | ESC: cancel",
                            new CvPoint(10, 24), HersheyFonts.HersheySimplex, 0.6, new Scalar(255, 255, 255), 2);
                        win.ShowImage(tmp);
                    }
                };

                win.ShowImage(img);
                Cv2.SetMouseCallback(winTitle, (ev, x, y, flags, ud) =>
                {
                    if (ev == MouseEventTypes.LButtonDown) { dragging = true; p0 = new CvPoint(x, y); cur = new Rect(); }
                    else if (ev == MouseEventTypes.MouseMove && dragging)
                    {
                        int w = Math.Abs(x - p0.X), h = Math.Abs(y - p0.Y);
                        int rx = Math.Min(x, p0.X), ry = Math.Min(y, p0.Y);
                        cur = SafeRoi(new Rect(rx, ry, Math.Max(1, w), Math.Max(1, h)), img.Size());
                        draw();
                    }
                    else if (ev == MouseEventTypes.LButtonUp)
                    {
                        dragging = false;
                        if (cur.Width > 3 && cur.Height > 3) { rois.Add(cur); cur = new Rect(); draw(); }
                    }
                }, IntPtr.Zero);

                draw();

                while (true)
                {
                    int key = Cv2.WaitKey(30);
                    if (key == 13 || key == 10) break;        // Enter
                    if (key == 27) { rois.Clear(); break; }   // ESC
                    if (key == 'r' || key == 'R' || key == 8) // R / Backspace
                    {
                        if (rois.Count > 0) rois.RemoveAt(rois.Count - 1);
                        draw();
                    }
                }
            }
            return rois;
        }

        // ---------------- 핵심 검사 로직 (다중 ROI + ROI별 설정) ----------------

        private void RunCheckCore(string goldenPath, string testPath)
        {
            var rois = LoadRoiListFromFile();
            if (rois.Count == 0) { Log("roi.txt가 비었거나 없습니다. 먼저 ROI 픽을 실행하세요."); return; }

            using (var golden = Cv2.ImRead(goldenPath, ImreadModes.Color))
            using (var test = Cv2.ImRead(testPath, ImreadModes.Color))
            {
                if (golden.Empty() || test.Empty()) { Log("이미지 읽기 실패"); return; }

                if (test.Size() != golden.Size()) Cv2.Resize(test, test, golden.Size());

                using (var testGray = ToGray(test))
                using (var gGray = ToGray(golden))
                using (var vis = test.Clone())
                {
                    bool[] roiPresent = new bool[rois.Count];
                    double scoreSum = 0;

                    for (int i = 0; i < rois.Count; i++)
                    {
                        var cfg = rois[i];
                        // ROI별 템플릿
                        Mat templGray = null;
                        string templPath = $"golden_roi_{cfg.Index:00}.png";
                        if (File.Exists(templPath)) templGray = Cv2.ImRead(templPath, ImreadModes.Grayscale);
                        if (templGray == null || templGray.Empty())
                        {
                            templGray?.Dispose();
                            templGray = new Mat(gGray, cfg.Roi).Clone();
                        }

                        // ROI별 검색영역 결정
                        Rect searchRect;
                        if (cfg.SearchRect.HasValue)
                        {
                            searchRect = SafeRoi(cfg.SearchRect.Value, testGray.Size());
                        }
                        else if (LOCK_TO_GOLDEN)
                        {
                            // 골든에서 그린 ROI 주변 작은 구역만 검색 → '움직임' 방지
                            searchRect = InflateAround(cfg.Roi, LOCK_MARGIN, testGray.Size());
                        }
                        else
                        {
                            bool useLeft = cfg.UseLeftSearch ?? USE_LEFT_SEARCH;
                            double leftRatio = cfg.LeftRatio ?? LEFT_WIDTH_RATIO;
                            searchRect = useLeft
                                ? new Rect(0, 0, (int)(testGray.Cols * leftRatio), testGray.Rows)
                                : new Rect(0, 0, testGray.Cols, testGray.Rows);
                        }

                        using (templGray)
                        using (var searchImg = new Mat(testGray, searchRect))
                        using (var searchPrep = PrepForMatch(searchImg))
                        using (var templPrep = PrepForMatch(templGray))
                        {
                            Rect found; double score; Mat scoreMap;

                            if (USE_MULTI_SCALE)
                                scoreMap = MultiScaleLocateTemplatePrepared(searchPrep, templGray, out found, out score);
                            else
                                scoreMap = LocateTemplate(searchPrep, templPrep, out found, out score);

                            found.X += searchRect.X; found.Y += searchRect.Y;
                            scoreSum += score;

                            if (scoreMap != null && !scoreMap.Empty())
                            {
                                Cv2.Normalize(scoreMap, scoreMap, 0, 255, NormTypes.MinMax);
                                scoreMap.ConvertTo(scoreMap, MatType.CV_8U);
                                Cv2.ApplyColorMap(scoreMap, scoreMap, ColormapTypes.Jet);
                                Cv2.ImWrite($"score_map_{cfg.Index:00}.png", scoreMap);
                                scoreMap.Dispose();
                            }

                            // ROI별 컷값(없으면 전역)
                            double sTH = cfg.ScoreTh ?? SCORE_TH;
                            double eTH = cfg.EdgeTh ?? EDGE_RATIO_TH;
                            double bTH = cfg.BrightTh ?? BRIGHT_RATIO_TH;

                            bool okByScore = score >= sTH;

                            // 보조 지표
                            var safeFound = SafeRoi(found, testGray.Size());
                            using (var testCrop = new Mat(testGray, safeFound).Clone())
                            using (var testEdges = MakeEdges(testCrop))
                            using (var gCrop = new Mat(gGray, cfg.Roi).Clone())
                            using (var gEdges = MakeEdges(gCrop))
                            {
                                Cv2.ImWrite($"test_roi_{cfg.Index:00}.png", testCrop);
                                Cv2.ImWrite($"test_edges_{cfg.Index:00}.png", testEdges);

                                double edgeG = Cv2.CountNonZero(gEdges);
                                double brightG = BrightRatio(gCrop);
                                double edgeT = Cv2.CountNonZero(testEdges);
                                double brightT = BrightRatio(testCrop);

                                bool edgeOk = (edgeG > 20) ? (edgeT >= eTH * edgeG) : true;
                                bool brightOk = (brightG > 0.01) ? (brightT >= bTH * brightG) : true;

                                bool present = okByScore && edgeOk && brightOk;
                                roiPresent[i] = present;

                                var col = present ? new Scalar(40, 220, 90) : new Scalar(40, 40, 255);
                                Cv2.Rectangle(vis, found, col, 2);
                                Cv2.PutText(vis, present ? $"P{cfg.Index}:OK" : $"P{cfg.Index}:NG",
                                    new CvPoint(found.X, Math.Max(20, found.Y - 6)),
                                    HersheyFonts.HersheySimplex, 0.6, col, 2);
                                Cv2.PutText(vis, $"s={score:0.00} (th={sTH:0.00})",
                                    new CvPoint(found.X, found.Y + found.Height + 18),
                                    HersheyFonts.HersheySimplex, 0.55, new Scalar(255, 255, 255), 1);
                            }
                        }
                    }

                    bool allOk = roiPresent.All(x => x);
                    bool anyOk = roiPresent.Any(x => x);
                    bool overall = ALL_ROIS_MUST_PRESENT ? allOk : anyOk;

                    var overallCol = overall ? new Scalar(40, 220, 90) : new Scalar(40, 40, 255);
                    Cv2.PutText(vis, overall ? "PRESENT" : "MISSING",
                        new CvPoint(12, 28), HersheyFonts.HersheySimplex, 1.0, overallCol, 2);

                    Cv2.ImWrite("result_vis.png", vis);
                    ShowMatOnPictureBox(vis);

                    Log($"[CHECK DONE] ROIs={rois.Count}  mode={(ALL_ROIS_MUST_PRESENT ? "ALL(AND)" : "ANY(OR)")}  " +
                        $"avgScore={scoreSum / rois.Count:0.00}  result={(overall ? "OK" : "NG")}");
                    Log("saved: result_vis.png, test_roi_XX.png, test_edges_XX.png, score_map_XX.png");
                }
            }
        }

        // ---------------- ROI 설정 파일 I/O ----------------

        private List<ROIConf> LoadRoiListFromFile()
        {
            var result = new List<ROIConf>();
            if (!File.Exists("roi.txt")) return result;

            int idx = 1;
            foreach (var raw in File.ReadAllLines("roi.txt"))
            {
                var line = raw.Trim();
                if (string.IsNullOrWhiteSpace(line)) continue;

                // "x y w h | key=val | key=val ..."
                var parts = line.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                var head = parts[0].Trim();
                var nums = head.Split(new[] { ' ', '\t', ',' }, StringSplitOptions.RemoveEmptyEntries);
                if (nums.Length < 4) continue;

                int x = int.Parse(nums[0]); int y = int.Parse(nums[1]);
                int w = int.Parse(nums[2]); int h = int.Parse(nums[3]);

                var cfg = new ROIConf { Index = idx, Roi = new Rect(x, y, w, h), Name = $"ROI{idx}" };

                for (int i = 1; i < parts.Length; i++)
                {
                    var seg = parts[i].Trim();
                    if (string.IsNullOrEmpty(seg)) continue;
                    var kv = seg.Split(new[] { '=' }, 2);
                    if (kv.Length < 2) continue;
                    var key = kv[0].Trim().ToLowerInvariant();
                    var val = kv[1].Trim();

                    double d; bool b;
                    switch (key)
                    {
                        case "name": cfg.Name = val; break;
                        case "score":
                        case "scoreth":
                            if (double.TryParse(val, NumberStyles.Any, CultureInfo.InvariantCulture, out d)) cfg.ScoreTh = d; break;
                        case "edge":
                        case "edgeth":
                        case "edgeratio":
                            if (double.TryParse(val, NumberStyles.Any, CultureInfo.InvariantCulture, out d)) cfg.EdgeTh = d; break;
                        case "bright":
                        case "brightth":
                            if (double.TryParse(val, NumberStyles.Any, CultureInfo.InvariantCulture, out d)) cfg.BrightTh = d; break;
                        case "left":
                        case "useleft":
                        case "leftsearch":
                            if (bool.TryParse(val, out b)) cfg.UseLeftSearch = b; break;
                        case "leftratio":
                            if (double.TryParse(val, NumberStyles.Any, CultureInfo.InvariantCulture, out d)) cfg.LeftRatio = d; break;
                        case "search":
                            {
                                var s = val.Split(new[] { ' ', '\t', ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
                                if (s.Length >= 4 &&
                                    int.TryParse(s[0], out var sx) &&
                                    int.TryParse(s[1], out var sy) &&
                                    int.TryParse(s[2], out var sw) &&
                                    int.TryParse(s[3], out var sh))
                                {
                                    cfg.SearchRect = new Rect(sx, sy, sw, sh);
                                }
                            }
                            break;
                    }
                }
                result.Add(cfg);
                idx++;
            }
            return result;
        }

        private void SaveRoiListToFile(List<ROIConf> list)
        {
            var lines = new List<string>();
            foreach (var c in list)
            {
                var tokens = new List<string>();
                tokens.Add($"{c.Roi.X} {c.Roi.Y} {c.Roi.Width} {c.Roi.Height}");
                if (!string.IsNullOrEmpty(c.Name)) tokens.Add($"name={c.Name}");
                if (c.ScoreTh.HasValue) tokens.Add($"score={c.ScoreTh.Value.ToString(CultureInfo.InvariantCulture)}");
                if (c.EdgeTh.HasValue) tokens.Add($"edge={c.EdgeTh.Value.ToString(CultureInfo.InvariantCulture)}");
                if (c.BrightTh.HasValue) tokens.Add($"bright={c.BrightTh.Value.ToString(CultureInfo.InvariantCulture)}");
                if (c.UseLeftSearch.HasValue) tokens.Add($"left={c.UseLeftSearch.Value}");
                if (c.LeftRatio.HasValue) tokens.Add($"leftratio={c.LeftRatio.Value.ToString(CultureInfo.InvariantCulture)}");
                if (c.SearchRect.HasValue)
                {
                    var r = c.SearchRect.Value;
                    tokens.Add($"search={r.X},{r.Y},{r.Width},{r.Height}");
                }
                lines.Add(string.Join(" | ", tokens));
            }
            File.WriteAllLines("roi.txt", lines);
        }

        // ---------------- 영상 유틸 ----------------

        private Mat LocateTemplate(Mat imgPrepared, Mat templPrepared, out Rect foundRect, out double maxVal)
        {
            var map = new Mat();
            Cv2.MatchTemplate(imgPrepared, templPrepared, map, TemplateMatchModes.CCoeffNormed);
            double minVal; CvPoint minLoc, maxLoc;
            Cv2.MinMaxLoc(map, out minVal, out maxVal, out minLoc, out maxLoc);
            foundRect = SafeRoi(new Rect(maxLoc.X, maxLoc.Y, templPrepared.Cols, templPrepared.Rows), imgPrepared.Size());
            return map;
        }

        private Mat MultiScaleLocateTemplatePrepared(Mat imgPrepared, Mat templGrayOriginal,
                                                     out Rect bestRect, out double bestVal)
        {
            bestVal = -1; bestRect = new Rect(0, 0, templGrayOriginal.Cols, templGrayOriginal.Rows);
            Mat bestMap = new Mat();

            for (double s = SCALE_MIN; s <= SCALE_MAX; s += SCALE_STEP)
            {
                using (var tGray = new Mat())
                {
                    Cv2.Resize(templGrayOriginal, tGray, new CvSize((int)(templGrayOriginal.Cols * s), (int)(templGrayOriginal.Rows * s)));
                    if (tGray.Cols < 10 || tGray.Rows < 10 || tGray.Cols >= imgPrepared.Cols || tGray.Rows >= imgPrepared.Rows) continue;

                    using (var tPrep = PrepForMatch(tGray))
                    using (var map = new Mat())
                    {
                        Cv2.MatchTemplate(imgPrepared, tPrep, map, TemplateMatchModes.CCoeffNormed);
                        double minV, maxV; CvPoint minL, maxL;
                        Cv2.MinMaxLoc(map, out minV, out maxV, out minL, out maxL);

                        if (maxV > bestVal)
                        {
                            bestVal = maxV;
                            bestRect = SafeRoi(new Rect(maxL.X, maxL.Y, tPrep.Cols, tPrep.Rows), imgPrepared.Size());
                            bestMap.Dispose();
                            bestMap = map.Clone();
                        }
                    }
                }
            }
            return bestMap;
        }

        private Mat ToGray(Mat bgr)
        {
            var g = new Mat();
            Cv2.CvtColor(bgr, g, ColorConversionCodes.BGR2GRAY);
            return g;
        }

        // 그래디언트(소벨) 강도 이미지 (0~255, u8)
        private Mat GradMagU8(Mat gray)
        {
            var gx16 = new Mat(); var gy16 = new Mat();
            Cv2.Sobel(gray, gx16, MatType.CV_16S, 1, 0, 3);
            Cv2.Sobel(gray, gy16, MatType.CV_16S, 0, 1, 3);

            var gx = new Mat(); var gy = new Mat();
            Cv2.ConvertScaleAbs(gx16, gx); Cv2.ConvertScaleAbs(gy16, gy);
            gx16.Dispose(); gy16.Dispose();

            var mag = new Mat();
            Cv2.AddWeighted(gx, 0.5, gy, 0.5, 0, mag);
            gx.Dispose(); gy.Dispose();

            Cv2.GaussianBlur(mag, mag, new CvSize(3, 3), 0);
            return mag;
        }

        // 매칭용 입력 준비: Edge Match면 소벨, 평평하면 평탄화로 폴백
        private Mat PrepForMatch(Mat gray)
        {
            if (USE_EDGE_MATCH)
            {
                var mag = GradMagU8(gray);
                int nz = Cv2.CountNonZero(mag);
                if (nz >= 30) return mag;      // 특징 충분
                mag.Dispose();                 // 너무 평평 → 폴백
                return ShadeNormalize(gray);
            }
            return ShadeNormalize(gray);
        }

        private Mat ShadeNormalize(Mat gray)
        {
            var bg = new Mat();
            Cv2.GaussianBlur(gray, bg, new CvSize(35, 35), 0);
            var norm = new Mat();
            Cv2.Divide(gray, bg, norm, scale: 255);
            bg.Dispose();
            return norm;
        }

        private Mat MakeEdges(Mat grayRoi)
        {
            var norm = ShadeNormalize(grayRoi);
            var blur = new Mat();
            Cv2.GaussianBlur(norm, blur, new CvSize(3, 3), 0);
            norm.Dispose();

            var edges = new Mat();
            Cv2.Canny(blur, edges, 30, 90, 3, true);
            blur.Dispose();

            var se = Cv2.GetStructuringElement(MorphShapes.Rect, new CvSize(3, 3));
            Cv2.MorphologyEx(edges, edges, MorphTypes.Open, se);
            Cv2.MorphologyEx(edges, edges, MorphTypes.Close, se);
            se.Dispose();
            return edges;
        }

        private double BrightRatio(Mat grayRoi)
        {
            var norm = ShadeNormalize(grayRoi);
            var bin = new Mat();
            Cv2.Threshold(norm, bin, 200, 255, ThresholdTypes.Binary);
            norm.Dispose();

            double r = Cv2.CountNonZero(bin) / (double)(bin.Rows * bin.Cols);
            bin.Dispose();
            return r;
        }

        private Rect SafeRoi(Rect r, CvSize sz)
        {
            int x = Math.Max(0, Math.Min(r.X, sz.Width - 1));
            int y = Math.Max(0, Math.Min(r.Y, sz.Height - 1));
            int w = Math.Max(1, Math.Min(r.Width, sz.Width - x));
            int h = Math.Max(1, Math.Min(r.Height, sz.Height - y));
            return new Rect(x, y, w, h);
        }

        // ---------------- ROI 설정 에디터 폼 ----------------

        private class ROIEditorForm : Form
        {
            private readonly DataGridView _grid = new DataGridView();
            private readonly Button _ok = new Button { Text = "저장", DialogResult = DialogResult.OK };
            private readonly Button _cancel = new Button { Text = "취소", DialogResult = DialogResult.Cancel };
            public List<ROIConf> Result { get; private set; }

            public ROIEditorForm(List<ROIConf> src)
            {
                this.Text = "ROI 설정";
                this.StartPosition = FormStartPosition.CenterParent;
                this.Size = new System.Drawing.Size(980, 480);

                _grid.Dock = DockStyle.Fill;
                _grid.AllowUserToAddRows = false;
                _grid.AutoGenerateColumns = false;
                _grid.RowHeadersVisible = false;

                void AddCol(string name, string header, Type t, bool readOnly = false, int w = 80)
                {
                    DataGridViewColumn col;
                    if (t == typeof(bool))
                        col = new DataGridViewCheckBoxColumn();
                    else
                        col = new DataGridViewTextBoxColumn();
                    col.Name = name; col.HeaderText = header; col.Width = w; col.ReadOnly = readOnly;
                    _grid.Columns.Add(col);
                }

                AddCol("Index", "#", typeof(int), true, 40);
                AddCol("Name", "Name", typeof(string), false, 90);

                AddCol("X", "X", typeof(int), true);
                AddCol("Y", "Y", typeof(int), true);
                AddCol("W", "W", typeof(int), true);
                AddCol("H", "H", typeof(int), true);

                AddCol("ScoreTh", "ScoreTH", typeof(double), false);
                AddCol("EdgeTh", "EdgeTH", typeof(double), false);
                AddCol("BrightTh", "BrightTH", typeof(double), false);

                AddCol("UseLeft", "Left", typeof(bool), false, 50);
                AddCol("LeftRatio", "LeftRatio", typeof(double), false, 80);

                AddCol("Sx", "SearchX", typeof(int), false);
                AddCol("Sy", "SearchY", typeof(int), false);
                AddCol("Sw", "SearchW", typeof(int), false);
                AddCol("Sh", "SearchH", typeof(int), false);

                foreach (var c in src)
                {
                    _grid.Rows.Add(
                        c.Index,
                        c.Name ?? $"ROI{c.Index}",
                        c.Roi.X, c.Roi.Y, c.Roi.Width, c.Roi.Height,
                        c.ScoreTh?.ToString(CultureInfo.InvariantCulture) ?? "",
                        c.EdgeTh?.ToString(CultureInfo.InvariantCulture) ?? "",
                        c.BrightTh?.ToString(CultureInfo.InvariantCulture) ?? "",
                        c.UseLeftSearch ?? false,
                        c.LeftRatio?.ToString(CultureInfo.InvariantCulture) ?? "",
                        c.SearchRect?.X, c.SearchRect?.Y, c.SearchRect?.Width, c.SearchRect?.Height
                    );
                }

                var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(8) };
                _ok.Width = 100; _cancel.Width = 100;
                buttons.Controls.Add(_ok); buttons.Controls.Add(_cancel);

                this.Controls.Add(_grid);
                this.Controls.Add(buttons);

                this.FormClosing += (s, e) =>
                {
                    if (this.DialogResult != DialogResult.OK) return;
                    try
                    {
                        Result = ParseGrid();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(this, ex.Message, "입력 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        e.Cancel = true;
                    }
                };
            }

            private List<ROIConf> ParseGrid()
            {
                var list = new List<ROIConf>();
                foreach (DataGridViewRow r in _grid.Rows)
                {
                    if (r.IsNewRow) continue;
                    int idx = Convert.ToInt32(r.Cells["Index"].Value);
                    string name = (r.Cells["Name"].Value ?? $"ROI{idx}").ToString();

                    int x = Convert.ToInt32(r.Cells["X"].Value);
                    int y = Convert.ToInt32(r.Cells["Y"].Value);
                    int w = Convert.ToInt32(r.Cells["W"].Value);
                    int h = Convert.ToInt32(r.Cells["H"].Value);

                    double? dNullable(string col)
                    {
                        var v = (r.Cells[col].Value ?? "").ToString().Trim();
                        if (string.IsNullOrEmpty(v)) return null;
                        if (double.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out var d)) return d;
                        throw new Exception($"{col} 입력이 잘못되었습니다. (row {idx})");
                    }

                    int? iNullable(string col)
                    {
                        var v = (r.Cells[col].Value ?? "").ToString().Trim();
                        if (string.IsNullOrEmpty(v)) return null;
                        if (int.TryParse(v, out var d)) return d;
                        throw new Exception($"{col} 입력이 잘못되었습니다. (row {idx})");
                    }

                    bool useLeft = false;
                    var vl = r.Cells["UseLeft"].Value;
                    if (vl != null && bool.TryParse(vl.ToString(), out var b)) useLeft = b;

                    var cfg = new ROIConf
                    {
                        Index = idx,
                        Name = name,
                        Roi = new Rect(x, y, w, h),
                        ScoreTh = dNullable("ScoreTh"),
                        EdgeTh = dNullable("EdgeTh"),
                        BrightTh = dNullable("BrightTh"),
                        UseLeftSearch = useLeft,
                        LeftRatio = dNullable("LeftRatio"),
                    };

                    var sx = iNullable("Sx"); var sy = iNullable("Sy"); var sw = iNullable("Sw"); var sh = iNullable("Sh");
                    if (sx.HasValue && sy.HasValue && sw.HasValue && sh.HasValue)
                        cfg.SearchRect = new Rect(sx.Value, sy.Value, sw.Value, sh.Value);

                    list.Add(cfg);
                }
                return list.OrderBy(c => c.Index).ToList();
            }
        }

        private void ToggleRoiEdit()
        {
            if (!_roiEditMode)
            {
                // === 편집 시작 ===
                ReadUiToOptions();

                // 경로 정리(공백/따옴표 제거)
                GOLDEN_PATH = (GOLDEN_PATH ?? "").Trim().Trim('"');

                // 골든 경로 없으면 자동으로 파일 선택창 띄우기
                if (!File.Exists(GOLDEN_PATH))
                {
                    Log($"Golden not found: {GOLDEN_PATH}");
                    using (var dlg = new OpenFileDialog())
                    {
                        dlg.Title = "골든 이미지 선택";
                        dlg.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp|All Files|*.*";
                        try
                        {
                            string init = GOLDEN_PATH;
                            if (File.Exists(init))
                                dlg.InitialDirectory = Path.GetDirectoryName(init);
                            else if (Directory.Exists(init))
                                dlg.InitialDirectory = init;
                            else if (!string.IsNullOrEmpty(init))
                            {
                                var dir = Path.GetDirectoryName(init);
                                if (Directory.Exists(dir)) dlg.InitialDirectory = dir;
                            }
                        }
                        catch { /* ignore */ }

                        if (dlg.ShowDialog(this.FindForm()) != DialogResult.OK) return; // 취소 시 편집 시작 안 함
                        tbGolden.Text = dlg.FileName;
                        GOLDEN_PATH = tbGolden.Text;
                        SavePrefs();
                    }
                }

                // 여기서부터는 골든이 반드시 존재
                if (!File.Exists(GOLDEN_PATH))
                {
                    MessageBox.Show("골든 이미지 경로가 올바르지 않습니다.", "ROI 편집",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 골든 로드
                _goldenBmp?.Dispose();
                _goldenMatColorForSave?.Dispose();
                using (var g = Cv2.ImRead(GOLDEN_PATH, ImreadModes.Color))
                {
                    if (g.Empty())
                    {
                        MessageBox.Show("골든 이미지를 읽지 못했습니다.", "ROI 편집",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    _goldenMatColorForSave = g.Clone();
                    _goldenBmp = OpenCvSharp.Extensions.BitmapConverter.ToBitmap(g);
                }

                // 편집 시작 시 기본 뷰 상태 초기화
                _zoom = 1f;
                _pan = new PointF(0, 0);

                // ▶ 빈 상태로 시작 / 강제 로드(ModifierKeys) 선택
                bool forceLoadExisting =
                    (Control.ModifierKeys & Keys.Control) == Keys.Control ||
                    (Control.ModifierKeys & Keys.Shift) == Keys.Shift;

                if (ROI_EDIT_STARTS_EMPTY && !forceLoadExisting)
                {
                    _roiEditing = new List<ROIConf>(); // 항상 새로 그리기
                    Log("ROI 편집: 빈 상태로 시작합니다. (기존 ROI를 불러오려면 Ctrl/Shift 누르고 버튼 클릭)");
                }
                else
                {
                    _roiEditing = LoadRoiListFromFile(); // 기존 roi.txt 로드
                    Log($"ROI 편집: 기존 ROI {_roiEditing.Count}개 불러옴");
                }

                _selSet.Clear();
                _lastSel = -1;
                _hit = HitType.None;
                _mouseDown = false;

                // 편집 모드에서는 직접 그리므로 PictureBox.Image는 비워 둠
                var old = pbView.Image;
                pbView.Image = null;
                old?.Dispose();

                _roiEditMode = true;
                btnPick.Text = "ROI 저장 종료";
                Log("ROI 편집: 드래그=추가, Shift+클릭=멀티선택, 핸들=리사이즈, Body=이동, Wheel=줌, 우클릭 드래그=팬, Del=삭제, Ctrl+C/V=복사/붙여넣기, ←↑→↓=미세이동(Shift×10), Enter=저장종료");
                pbView.Focus();
                pbView.Invalidate();
            }
            else
            {
                // === 저장 & 편집 종료 ===
                try
                {
                    if (_goldenMatColorForSave != null && !_goldenMatColorForSave.Empty())
                    {
                        using (var gGray = ToGray(_goldenMatColorForSave))
                        {
                            int idx = 1;
                            foreach (var c in _roiEditing)
                            {
                                var r = SafeRoi(c.Roi, gGray.Size());
                                using (var t = new Mat(gGray, r))
                                    Cv2.ImWrite($"golden_roi_{idx:00}.png", t);

                                c.Index = idx;
                                if (string.IsNullOrWhiteSpace(c.Name))
                                    c.Name = $"ROI{idx}";
                                idx++;
                            }
                        }
                    }

                    SaveRoiListToFile(_roiEditing);
                    Log($"ROI {_roiEditing.Count}개 저장 완료.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("ROI 저장 중 오류: " + ex.Message, "ROI 편집",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    Log("ERROR while saving ROIs: " + ex);
                }
                finally
                {
                    _roiEditMode = false;
                    btnPick.Text = "ROI 편집";

                    // 편집 종료 후 화면 갱신(검사 이미지/결과를 다시 그리고 싶다면 외부에서 ShowMatOnPictureBox 호출)
                    pbView.Invalidate();
                }
            }
        }


        // PictureBox 안에서 이미지가 차지하는 사각형(레터박스 고려)
        // 편집 모드일 때: 우리가 직접 그리는 이미지 사각형(줌/팬 반영)
        private RectangleF ViewImgRect()
        {
            if (_roiEditMode && _goldenBmp != null)
            {
                return new RectangleF(-_pan.X * _zoom, -_pan.Y * _zoom,
                                      _goldenBmp.Width * _zoom, _goldenBmp.Height * _zoom);
            }
            // 편집 모드가 아닐 땐 기존 레터박스 로직
            if (pbView.Image == null) return RectangleF.Empty;
            float cw = pbView.ClientSize.Width, ch = pbView.ClientSize.Height;
            float iw = pbView.Image.Width, ih = pbView.Image.Height;
            float ca = cw / ch, ia = iw / ih;
            float vw, vh, vx, vy;
            if (ia >= ca) { vw = cw; vh = vw / ia; vx = 0; vy = (ch - vh) / 2f; }
            else { vh = ch; vw = vh * ia; vx = (cw - vw) / 2f; vy = 0; }
            return new RectangleF(vx, vy, vw, vh);
        }

        private SDRectF RectImgToView(OpenCvSharp.Rect r)
        {
            var p1 = ImgToView(new CvPoint(r.X, r.Y));
            var p2 = ImgToView(new CvPoint(r.X + r.Width, r.Y + r.Height));
            return new SDRectF(p1.X, p1.Y, p2.X - p1.X, p2.Y - p1.Y);
        }

        private CvPoint ViewToImg(SDPoint p)
        {
            if (_roiEditMode && _goldenBmp != null)
            {
                double ix = p.X / _zoom + _pan.X;
                double iy = p.Y / _zoom + _pan.Y;
                int w = _goldenBmp.Width, h = _goldenBmp.Height;
                int x = (int)Math.Round(Math.Max(0, Math.Min(w - 1, ix)));
                int y = (int)Math.Round(Math.Max(0, Math.Min(h - 1, iy)));
                return new CvPoint(x, y);
            }
            var r = ViewImgRect();
            if (pbView.Image == null || r.IsEmpty) return new CvPoint(0, 0);
            double x2 = (p.X - r.X) * pbView.Image.Width / r.Width;
            double y2 = (p.Y - r.Y) * pbView.Image.Height / r.Height;
            int ix2 = (int)Math.Round(Math.Max(0, Math.Min(pbView.Image.Width - 1, x2)));
            int iy2 = (int)Math.Round(Math.Max(0, Math.Min(pbView.Image.Height - 1, y2)));
            return new CvPoint(ix2, iy2);
        }

        private SDPointF ImgToView(CvPoint p)
        {
            if (_roiEditMode && _goldenBmp != null)
                return new SDPointF((p.X - _pan.X) * _zoom, (p.Y - _pan.Y) * _zoom);

            var r = ViewImgRect();
            if (pbView.Image == null || r.IsEmpty) return SDPointF.Empty;
            return new SDPointF(r.X + p.X * r.Width / pbView.Image.Width,
                                r.Y + p.Y * r.Height / pbView.Image.Height);
        }

        private void PbView_Paint(object sender, PaintEventArgs e)
        {
            if (_roiEditMode)
            {
                // 1) 이미지 직접 그림
                if (_goldenBmp != null)
                {
                    e.Graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    var dst = ViewImgRect();
                    e.Graphics.DrawImage(_goldenBmp, dst);
                }

                // 2) ROI 오버레이
                if (_roiEditing != null)
                {
                    for (int i = 0; i < _roiEditing.Count; i++)
                    {
                        var r = RectImgToView(_roiEditing[i].Roi);
                        bool sel = _selSet.Contains(i);
                        var col = sel ? Color.Lime : Color.FromArgb(0, 220, 120);

                        using (var pen = new Pen(col, sel ? 2.5f : 2f))
                            e.Graphics.DrawRectangle(pen, r.X, r.Y, r.Width, r.Height);

                        using (var f = new Font("Segoe UI", 9f, FontStyle.Bold))
                        using (var br = new SolidBrush(col))
                            e.Graphics.DrawString(_roiEditing[i].Name ?? $"ROI{i + 1}", f, br, r.X + 2, r.Y - 16);
                    }

                    // 단일 선택 시 핸들
                    if (_selSet.Count == 1)
                    {
                        var r = RectImgToView(_roiEditing[_selSet[0]].Roi);
                        foreach (var pt in HandlePoints(r))
                            e.Graphics.FillRectangle(Brushes.White, pt.X - HANDLE / 2f, pt.Y - HANDLE / 2f, HANDLE, HANDLE);
                    }

                    // 다중 선택 시 바운딩 박스
                    if (_selSet.Count > 1)
                    {
                        var vr = GetSelectionViewBounds();
                        using (var pen = new Pen(Color.Lime, 2f) { DashStyle = DashStyle.Dash })
                            e.Graphics.DrawRectangle(pen, vr.X, vr.Y, vr.Width, vr.Height);
                    }
                }
            }
        }

        private RectangleF GetSelectionViewBounds()
        {
            var boxes = _selSet.Select(i => RectImgToView(_roiEditing[i].Roi)).ToList();
            if (boxes.Count == 0) return RectangleF.Empty;

            float left = boxes.Min(b => b.Left);
            float top = boxes.Min(b => b.Top);
            float right = boxes.Max(b => b.Right);
            float bottom = boxes.Max(b => b.Bottom);

            return new RectangleF(left, top, right - left, bottom - top);
        }


        private IEnumerable<PointF> HandlePoints(RectangleF r)
        {
            yield return new PointF(r.Left, r.Top);    // NW
            yield return new PointF(r.Right, r.Top);    // NE
            yield return new PointF(r.Left, r.Bottom); // SW
            yield return new PointF(r.Right, r.Bottom); // SE
            yield return new PointF(r.Left + r.Width / 2f, r.Top);    // N
            yield return new PointF(r.Left + r.Width / 2f, r.Bottom); // S
            yield return new PointF(r.Left, r.Top + r.Height / 2f);  // W
            yield return new PointF(r.Right, r.Top + r.Height / 2f);  // E
        }

        private bool Near(System.Drawing.Point p, PointF q)
        {
            return Math.Abs(p.X - q.X) <= HANDLE && Math.Abs(p.Y - q.Y) <= HANDLE;
        }

        private Cursor CursorFromHit(HitType h)
        {
            switch (h)
            {
                case HitType.N:
                case HitType.S: return Cursors.SizeNS;
                case HitType.E:
                case HitType.W: return Cursors.SizeWE;
                case HitType.NE:
                case HitType.SW: return Cursors.SizeNESW;
                case HitType.NW:
                case HitType.SE: return Cursors.SizeNWSE;
                case HitType.Body: return Cursors.SizeAll;
                default: return Cursors.Cross;
            }
        }

        private void PbView_MouseDown(object sender, MouseEventArgs e)
        {
            if (!_roiEditMode || _goldenBmp == null) return;

            // 팬(우클릭/휠 클릭)
            if (e.Button == MouseButtons.Right || e.Button == MouseButtons.Middle)
            {
                _panning = true;
                _panStartView = e.Location;
                _panStartImg = _pan;
                return;
            }

            _mouseDown = true;
            var hit = HitTestAll(e.Location);

            if (hit.idx >= 0)
            {
                // 선택 갱신 (Shift로 토글)
                if ((ModifierKeys & Keys.Shift) == Keys.Shift)
                {
                    if (_selSet.Contains(hit.idx)) _selSet.Remove(hit.idx);
                    else _selSet.Add(hit.idx);
                }
                else
                {
                    _selSet.Clear(); _selSet.Add(hit.idx);
                }
                _lastSel = hit.idx; _hit = hit.hit;
                _dragStartImg = ViewToImg(e.Location);
                _dragOrigRect = _roiEditing[hit.idx].Roi;

                if (_hit == HitType.Body && _selSet.Count > 1)
                    _dragOrigRects = _selSet.ToDictionary(i => i, i => _roiEditing[i].Roi);
                else
                    _dragOrigRects = null;
            }
            else
            {
                // 빈 공간: Shift 없으면 선택 해제, 새 ROI 생성 시작
                if ((ModifierKeys & Keys.Shift) != Keys.Shift) { _selSet.Clear(); _lastSel = -1; }
                var p = ViewToImg(e.Location);
                var rect = new Rect(p.X, p.Y, 1, 1);
                var cfg = new ROIConf { Index = _roiEditing.Count + 1, Roi = rect, Name = $"ROI{_roiEditing.Count + 1}" };
                _roiEditing.Add(cfg);
                _selSet.Clear(); _selSet.Add(_roiEditing.Count - 1);
                _lastSel = _selSet[0];
                _hit = HitType.SE; // 오른쪽/아래로 드래그
                _dragStartImg = p;
                _dragOrigRect = rect;
            }
            pbView.Invalidate();
        }

        private void PbView_MouseMove(object sender, MouseEventArgs e)
        {
            if (!_roiEditMode || _goldenBmp == null)
            { pbView.Cursor = Cursors.Default; return; }

            // 팬
            if (_panning)
            {
                var dvx = e.Location.X - _panStartView.X;
                var dvy = e.Location.Y - _panStartView.Y;
                _pan = new PointF(_panStartImg.X - dvx / _zoom, _panStartImg.Y - dvy / _zoom);
                pbView.Invalidate(); return;
            }

            if (!_mouseDown)
            {
                var ht = HitTestAll(e.Location);
                pbView.Cursor = CursorFromHit(ht.hit);
                return;
            }

            if (_selSet.Count == 0) return;

            var cur = ViewToImg(e.Location);
            int dx = cur.X - _dragStartImg.X;
            int dy = cur.Y - _dragStartImg.Y;
            const int MIN = 6;

            if (_hit == HitType.Body && _selSet.Count > 1 && _dragOrigRects != null)
            {
                foreach (var kv in _dragOrigRects.ToList())
                {
                    var r0 = kv.Value;
                    var r = new Rect(r0.X + dx, r0.Y + dy, r0.Width, r0.Height);
                    r = SafeRoi(r, new OpenCvSharp.Size(_goldenBmp.Width, _goldenBmp.Height));
                    _roiEditing[kv.Key].Roi = r;
                }
            }
            else
            {
                int idx = _selSet[0];
                var r = _dragOrigRect;
                switch (_hit)
                {
                    case HitType.Body: r.X += dx; r.Y += dy; break;
                    case HitType.N: r.Y += dy; r.Height -= dy; break;
                    case HitType.S: r.Height += dy; break;
                    case HitType.W: r.X += dx; r.Width -= dx; break;
                    case HitType.E: r.Width += dx; break;
                    case HitType.NW: r.X += dx; r.Width -= dx; r.Y += dy; r.Height -= dy; break;
                    case HitType.NE: r.Width += dx; r.Y += dy; r.Height -= dy; break;
                    case HitType.SW: r.X += dx; r.Width -= dx; r.Height += dy; break;
                    case HitType.SE: r.Width += dx; r.Height += dy; break;
                }
                if (r.Width < MIN) r.Width = MIN;
                if (r.Height < MIN) r.Height = MIN;
                r = SafeRoi(r, new OpenCvSharp.Size(_goldenBmp.Width, _goldenBmp.Height));
                _roiEditing[idx].Roi = r;
            }
            pbView.Invalidate();
        }

        private void PbView_MouseUp(object sender, MouseEventArgs e)
        {
            if (!_roiEditMode) return;
            _mouseDown = false;
            _panning = false;
            _dragOrigRects = null;
        }

        private void PbView_MouseWheel(object sender, MouseEventArgs e)
        {
            if (!_roiEditMode || _goldenBmp == null) return;
            float old = _zoom;
            float factor = (e.Delta > 0) ? 1.1f : (1f / 1.1f);
            float nz = Math.Max(0.1f, Math.Min(20f, old * factor));
            var imgPt = ViewToImg(e.Location); // 줌 고정점
            _zoom = nz;
            _pan = new PointF((float)(imgPt.X - e.X / _zoom), (float)(imgPt.Y - e.Y / _zoom));
            pbView.Invalidate();
        }

        private void PbView_KeyDown(object sender, KeyEventArgs e)
        {
            if (!_roiEditMode) return;

            // 복사/붙여넣기
            if (e.Control && e.KeyCode == Keys.C)
            {
                _clipboard = _selSet.Select(i => _roiEditing[i].Roi).ToList();
                e.Handled = true; return;
            }
            if (e.Control && e.KeyCode == Keys.V && _clipboard.Count > 0)
            {
                int next = _roiEditing.Count + 1;
                _selSet.Clear();
                foreach (var r0 in _clipboard)
                {
                    var r = new Rect(r0.X + 10, r0.Y + 10, r0.Width, r0.Height);
                    r = SafeRoi(r, new OpenCvSharp.Size(_goldenBmp.Width, _goldenBmp.Height));
                    _roiEditing.Add(new ROIConf { Index = next, Roi = r, Name = $"ROI{next}" });
                    _selSet.Add(_roiEditing.Count - 1);
                    next++;
                }
                _lastSel = _selSet.Last();
                pbView.Invalidate();
                e.Handled = true; return;
            }

            // 삭제
            if (e.KeyCode == Keys.Delete && _selSet.Count > 0)
            {
                // 뒤에서부터 지우기
                foreach (var i in _selSet.OrderByDescending(x => x))
                    _roiEditing.RemoveAt(i);
                for (int i = 0; i < _roiEditing.Count; i++) _roiEditing[i].Index = i + 1;
                _selSet.Clear(); _lastSel = -1;
                pbView.Invalidate();
                e.Handled = true; return;
            }

            // 미세 이동
            if (_selSet.Count > 0 && (e.KeyCode == Keys.Left || e.KeyCode == Keys.Right || e.KeyCode == Keys.Up || e.KeyCode == Keys.Down))
            {
                int step = e.Shift ? 10 : 1; // Shift ×10
                int dx = 0, dy = 0;
                if (e.KeyCode == Keys.Left) dx = -step;
                if (e.KeyCode == Keys.Right) dx = step;
                if (e.KeyCode == Keys.Up) dy = -step;
                if (e.KeyCode == Keys.Down) dy = step;

                foreach (var i in _selSet)
                {
                    var r = _roiEditing[i].Roi;
                    r.X += dx; r.Y += dy;
                    r = SafeRoi(r, new OpenCvSharp.Size(_goldenBmp.Width, _goldenBmp.Height));
                    _roiEditing[i].Roi = r;
                }
                pbView.Invalidate();
                e.Handled = true; return;
            }

            // 엔터: 저장 종료
            if (e.KeyCode == Keys.Enter) { ToggleRoiEdit(); e.Handled = true; }
        }

        private (int idx, HitType hit) HitTestAll(SDPoint pt)
        {
            for (int i = _roiEditing.Count - 1; i >= 0; i--)
            {
                var r = RectImgToView(_roiEditing[i].Roi);
                var corners = HandlePoints(r).ToList();
                if (Near(pt, corners[0])) return (i, HitType.NW);
                if (Near(pt, corners[1])) return (i, HitType.NE);
                if (Near(pt, corners[2])) return (i, HitType.SW);
                if (Near(pt, corners[3])) return (i, HitType.SE);
                if (Near(pt, corners[4])) return (i, HitType.N);
                if (Near(pt, corners[5])) return (i, HitType.S);
                if (Near(pt, corners[6])) return (i, HitType.W);
                if (Near(pt, corners[7])) return (i, HitType.E);
                if (r.Contains(pt)) return (i, HitType.Body);
            }
            return (-1, HitType.None);
        }

        private Rect InflateAround(Rect roi, double margin, OpenCvSharp.Size imgSize)
        {
            // roi 중심을 기준으로 margin 비율만큼 사방으로 확장
            int cx = roi.X + roi.Width / 2;
            int cy = roi.Y + roi.Height / 2;

            int w = (int)Math.Round(roi.Width * (1.0 + 2.0 * margin));
            int h = (int)Math.Round(roi.Height * (1.0 + 2.0 * margin));

            int x = cx - w / 2;
            int y = cy - h / 2;

            var r = new Rect(x, y, w, h);
            return SafeRoi(r, imgSize);
        }


        // ---------------- 끝 ----------------
    }
}