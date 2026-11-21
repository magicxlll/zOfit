using System.Drawing;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ZOfit.Engine.Generators;
using ZOfit.Engine.Models;
using ZOfit.Engine.Parsers;

namespace ZOfit.Engine.UI;

public sealed class RebarEditorForm : Form
{
    [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hWnd);
    private readonly EditRebarUiRequest _req;
    private readonly List<RebarConfigRowDto> _bars = new();
    private readonly List<RebarCalcPreview> _calcs = new();
    private bool _suppressGrid;

    private readonly Panel _diagramPanel = new();
    private readonly Panel _shapePanel = new();
    private readonly DataGridView _grid = new();
    private readonly Label _lblShapeDetail = new();
    private readonly Label _lblTotalKg = new();
    private readonly FlowLayoutPanel _quickFlow = new();

    public EditRebarUiResponse Result { get; private set; }

    public RebarEditorForm(EditRebarUiRequest req)
    {
        _req = req;
        Result = new EditRebarUiResponse { Accepted = false, ElementId = req.Element.ElementId };

        if (req.ExistingRows != null && req.ExistingRows.Count > 0)
        {
            foreach (var r in req.ExistingRows)
            {
                if (!string.Equals(r.RowKind, "BLK_HDR", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(r.RowKind, "BLK_END", StringComparison.OrdinalIgnoreCase))
                {
                    _bars.Add(CloneRow(r));
                }
            }
        }
        else if (!string.IsNullOrWhiteSpace(req.Element.RebarTemplate) &&
                 !string.Equals(req.Element.RebarTemplate.Trim(), "NONE", StringComparison.OrdinalIgnoreCase))
        {
            _bars.AddRange(new RebarBbsGenerator().ExpandSingleElementToConfigRows(req.Element));
        }

        if (_bars.Count == 0)
        {
            _bars.Add(RebarParamCalculator.CreateQuickRoleRow(req.Element, req.Spec, DefaultQuickKey(req.Element.Group), 1));
        }

        BuildUi();
        RecomputeAll();
        PopulateGrid();
        Load += (_, _) =>
        {
            // Force visible even when parent (Excel WshShell.Run 0) launched process with STARTF_USESHOWWINDOW=SW_HIDE
            ShowWindow(Handle, 5); // SW_SHOW
            BringToFront();
            Activate();
            SetForegroundWindow(Handle);
        };
    }

    private static string DefaultQuickKey(string? group) => (group ?? "").Trim().ToLowerInvariant() switch
    {
        "beam" or "groundbeam" or "lintel" => "TOP",
        "column" => "MAIN.A",
        "wall" => "MAIN.V",
        "stair" => "ST-01",
        _ => "BOT.X"
    };

    private static RebarConfigRowDto CloneRow(RebarConfigRowDto s) => new()
    {
        ConfigRow = s.ConfigRow, RowKind = "BAR", ElementId = s.ElementId,
        Building = s.Building, Group = s.Group, BarId = s.BarId, Role = s.Role,
        ShapeCode = string.IsNullOrWhiteSpace(s.ShapeCode) ? "D01" : s.ShapeCode,
        Dia = s.Dia > 0 ? s.Dia : 14, LayerMult = s.LayerMult > 0 ? s.LayerMult : 1,
        BarDir = string.IsNullOrWhiteSpace(s.BarDir) ? "X" : s.BarDir,
        DistDir = string.IsNullOrWhiteSpace(s.DistDir) ? "FIXED" : s.DistDir,
        HookLExpr = s.HookLExpr, CoverExpr = s.CoverExpr, HookRExpr = s.HookRExpr,
        Space1Mm = s.Space1Mm, Space2Mm = s.Space2Mm,
        OverrideQtyExpr = s.OverrideQtyExpr, OverrideLenExpr = s.OverrideLenExpr,
        DeltaDistExpr = s.DeltaDistExpr, DeltaLenExpr = s.DeltaLenExpr,
        Le1aMm = s.Le1aMm, Le1bMm = s.Le1bMm, Ri2bMm = s.Ri2bMm, Ri2aMm = s.Ri2aMm,
        Note = s.Note
    };

    private void BuildUi()
    {
        var el = _req.Element;
        Text = $"zOfit QS Pro - Nhập Thép Trực Quan: {el.Mark} ({el.ElementId})";
        Size = new Size(1240, 770);
        MinimumSize = new Size(1020, 650);
        StartPosition = FormStartPosition.CenterParent;
        Font = new Font("Segoe UI", 9f);
        BackColor = Color.FromArgb(245, 247, 250);

        // 1. Top Element Banner
        var topBanner = new Panel { Dock = DockStyle.Top, Height = 54, BackColor = Color.FromArgb(22, 54, 92), Padding = new Padding(14, 8, 14, 8) };
        var lblTitle = new Label
        {
            Dock = DockStyle.Left, AutoSize = true, ForeColor = Color.White,
            Font = new Font("Segoe UI", 11.5f, FontStyle.Bold),
            Text = $"{el.Mark}  [{el.Group}]  —  {el.Description}"
        };
        var lblSource = new Label
        {
            Dock = DockStyle.Left, AutoSize = true, ForeColor = Color.FromArgb(144, 238, 144),
            Font = new Font("Segoe UI", 9f, FontStyle.Bold), Padding = new Padding(15, 3, 0, 0),
            Text = "● Nguồn: Tự động từ AutoCAD MCP (Port 9877) - Không cần nhập tay"
        };
        var lblDim = new Label
        {
            Dock = DockStyle.Right, AutoSize = true, ForeColor = Color.FromArgb(255, 232, 150),
            Font = new Font("Segoe UI", 10f, FontStyle.Bold), TextAlign = ContentAlignment.MiddleRight,
            Text = $"SL: {el.Qty:0.##} cái   |   X={el.LengthM * 1000:0} × Y={el.WidthM * 1000:0} × H={el.HeightM * 1000:0} mm   |   Preset: {(string.IsNullOrWhiteSpace(el.RebarTemplate) ? "TÙY BIẾN" : el.RebarTemplate)}"
        };
        topBanner.Controls.Add(lblDim);
        topBanner.Controls.Add(lblSource);
        topBanner.Controls.Add(lblTitle);

        // 2. Bottom Bar
        var bottomBar = new Panel { Dock = DockStyle.Bottom, Height = 54, BackColor = Color.FromArgb(236, 241, 248), Padding = new Padding(14, 9, 14, 9) };
        _lblTotalKg.Dock = DockStyle.Left; _lblTotalKg.AutoSize = true;
        _lblTotalKg.Font = new Font("Segoe UI", 10.5f, FontStyle.Bold);
        _lblTotalKg.ForeColor = Color.FromArgb(22, 54, 92);

        var btnCancel = new Button { Text = "Hủy bỏ", Width = 95, Height = 34, Dock = DockStyle.Right, FlatStyle = FlatStyle.System };
        btnCancel.Click += (_, _) => { Result = new EditRebarUiResponse { Accepted = false, ElementId = el.ElementId }; DialogResult = DialogResult.Cancel; Close(); };

        var spacer = new Panel { Width = 10, Dock = DockStyle.Right };
        var btnOk = new Button
        {
            Text = "✓ Sử dụng & Ghi vào Excel", Width = 210, Height = 34, Dock = DockStyle.Right,
            BackColor = Color.FromArgb(31, 78, 120), ForeColor = Color.White, FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9.5f, FontStyle.Bold)
        };
        btnOk.FlatAppearance.BorderSize = 0;
        btnOk.Click += (_, _) => OnAccept();

        var spacerCad = new Panel { Width = 10, Dock = DockStyle.Right };
        var btnRescanCad = new Button
        {
            Text = "🖱️ Quét vùng khác trên CAD", Width = 190, Height = 34, Dock = DockStyle.Right,
            BackColor = Color.FromArgb(47, 85, 151), ForeColor = Color.White, FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9f, FontStyle.Bold)
        };
        btnRescanCad.FlatAppearance.BorderSize = 0;
        btnRescanCad.Click += async (_, _) =>
        {
            WindowState = FormWindowState.Minimized;
            var client = new AutoCadMcpClient();

            var dr = MessageBox.Show(
                "Vui lòng chuyển sang AutoCAD, dùng chuột kéo chọn (bôi đen) bảng chi tiết dầm cần bóc tách.\n\nSau khi đã quét chọn xong trên CAD, bấm 'OK' để zOfit đọc dữ liệu!",
                "zOfit QS Pro - Quét Mặt Cắt Dầm từ AutoCAD",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Information);

            if (dr == DialogResult.OK)
            {
                try
                {
                    var ents = await client.GetSelectionAsync();
                    if (ents.Count == 0)
                    {
                        MessageBox.Show(
                            "Chưa có đối tượng nào được chọn trên AutoCAD!\nVui lòng bôi đen trọn vẹn bảng mặt cắt dầm trên CAD trước rồi bấm lại.",
                            "Chưa chọn đối tượng",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                        WindowState = FormWindowState.Normal;
                        BringToFront();
                        Activate();
                        return;
                    }

                    var parser = new CadBeamSectionParser();
                    var resp = parser.Parse(new CadBeamSectionParseRequest
                    {
                        Building = _req.Element.Building,
                        Floor = "TANG1",
                        DefaultLengthM = _req.Element.LengthM > 0 ? _req.Element.LengthM : 6.0,
                        Entities = ents
                    });

                    if (resp.BeamCount > 0)
                    {
                        var b = resp.Beams[0];
                        _req.Element = b.Element;
                        _bars.Clear();
                        _bars.AddRange(b.RebarRows);
                        lblTitle.Text = $"{b.Element.Mark}  [{b.Element.Group}]  —  {b.Element.Description}";
                        lblDim.Text = $"SL: {b.Element.Qty:0.##} cái   |   X={b.Element.LengthM * 1000:0} × Y={b.Element.WidthM * 1000:0} × H={b.Element.HeightM * 1000:0} mm   |   Preset: BEAM_PARAMETRIC";
                        RecomputeAll();
                        PopulateGrid();
                        Text = $"zOfit QS Pro - Đã Quét Tự Động từ CAD: {b.Element.Mark} ({b.Element.WidthM * 1000:0}x{b.Element.HeightM * 1000:0})";
                        _diagramPanel.Invalidate();
                        _shapePanel.Invalidate();
                        MessageBox.Show(
                            $"Đã quét thành công {resp.BeamCount} dầm ({string.Join(", ", resp.Beams.Select(x => x.Element.Mark))}) từ AutoCAD!\nSố liệu và hình vẽ đã được cập nhật tự động.",
                            "Quét CAD Thành Công",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show(
                            "Không nhận diện được bảng mặt cắt dầm nào trong vùng vừa chọn trên AutoCAD.\nVui lòng bôi đen trọn vẹn cả bảng mặt cắt dầm (gồm tên dầm, mặt cắt, thép trên, dưới, đai).",
                            "Không tìm thấy dầm",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        $"Lỗi khi kết nối AutoCAD MCP (Port 9877): {ex.Message}\nHãy kiểm tra xem AutoCAD đã kết nối MCP chưa.",
                        "Lỗi kết nối",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }

            WindowState = FormWindowState.Normal;
            BringToFront();
            Activate();
        };

        bottomBar.Controls.Add(_lblTotalKg);
        bottomBar.Controls.Add(btnRescanCad);
        bottomBar.Controls.Add(spacerCad);
        bottomBar.Controls.Add(btnOk);
        bottomBar.Controls.Add(spacer);
        bottomBar.Controls.Add(btnCancel);

        // 3. Upper Visual Area (Left Quick Roles | Center GDI+ Diagram | Right Shape Inspector)
        var upperPanel = new Panel { Dock = DockStyle.Top, Height = 305, Padding = new Padding(8) };

        var grpQuick = new GroupBox
        {
            Text = "Thêm nhanh theo vị trí", Dock = DockStyle.Left, Width = 230,
            Font = new Font("Segoe UI", 9f, FontStyle.Bold), ForeColor = Color.FromArgb(22, 54, 92)
        };
        _quickFlow.Dock = DockStyle.Fill; _quickFlow.AutoScroll = true;
        _quickFlow.FlowDirection = FlowDirection.LeftToRight; _quickFlow.Padding = new Padding(4);
        BuildQuickRoleButtons(el.Group);
        grpQuick.Controls.Add(_quickFlow);

        var grpShape = new GroupBox
        {
            Text = "Chi tiết thanh đang chọn", Dock = DockStyle.Right, Width = 245,
            Font = new Font("Segoe UI", 9f, FontStyle.Bold), ForeColor = Color.FromArgb(22, 54, 92)
        };
        typeof(Panel).GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)?.SetValue(_shapePanel, true);
        _shapePanel.Dock = DockStyle.Top; _shapePanel.Height = 145; _shapePanel.BorderStyle = BorderStyle.FixedSingle;
        _shapePanel.Paint += (_, e) =>
        {
            int idx = SelectedRowIdx();
            StructuralDiagramRenderer.RenderShapePreview(e.Graphics, _shapePanel.ClientRectangle,
                idx >= 0 && idx < _bars.Count ? _bars[idx] : null,
                idx >= 0 && idx < _calcs.Count ? _calcs[idx] : null);
        };
        _lblShapeDetail.Dock = DockStyle.Fill; _lblShapeDetail.Font = new Font("Segoe UI", 8.5f);
        _lblShapeDetail.ForeColor = Color.FromArgb(40, 40, 40); _lblShapeDetail.Padding = new Padding(6);
        grpShape.Controls.Add(_lblShapeDetail);
        grpShape.Controls.Add(_shapePanel);

        var grpDiag = new GroupBox
        {
            Text = "Bản vẽ Cấu kiện & Chú thích Cốt thép (Chọn dòng bên dưới để tô sáng thanh tương ứng)",
            Dock = DockStyle.Fill, Font = new Font("Segoe UI", 9f, FontStyle.Bold), ForeColor = Color.FromArgb(22, 54, 92)
        };
        typeof(Panel).GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)?.SetValue(_diagramPanel, true);
        _diagramPanel.Dock = DockStyle.Fill; _diagramPanel.BorderStyle = BorderStyle.FixedSingle;
        _diagramPanel.Paint += (_, e) => StructuralDiagramRenderer.RenderElementDiagram(e.Graphics, _diagramPanel.ClientRectangle, _req.Element, _bars, _calcs, SelectedRowIdx());
        grpDiag.Controls.Add(_diagramPanel);

        upperPanel.Controls.Add(grpDiag);
        upperPanel.Controls.Add(grpQuick);
        upperPanel.Controls.Add(grpShape);

        // 4. Grid Toolbar + Editable Grid
        var midWrap = new Panel { Dock = DockStyle.Fill, Padding = new Padding(8, 0, 8, 6) };
        var barOps = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 36, BackColor = Color.FromArgb(237, 242, 248), Padding = new Padding(4) };
        AddToolBtn(barOps, "Nhân bản dòng", () => DuplicateSelected());
        AddToolBtn(barOps, "Lên", () => MoveSelected(-1));
        AddToolBtn(barOps, "Xuống", () => MoveSelected(1));
        AddToolBtn(barOps, "Xóa dòng", () => DeleteSelected());
        AddToolBtn(barOps, "Nạp lại Mẫu chuẩn", () => ReloadPreset());
        var lblTip = new Label
        {
            AutoSize = true, Margin = new Padding(14, 7, 4, 0), ForeColor = Color.FromArgb(90, 90, 90), Font = new Font("Segoe UI", 8.5f, FontStyle.Italic),
            Text = "Mẹo: Mỏ <50 tự nhân ×D (vd 20=20D), ≥50 lấy thẳng mm | Để trống SL.m/L.m = tự tính"
        };
        barOps.Controls.Add(lblTip);

        ConfigureGrid();
        midWrap.Controls.Add(_grid);
        midWrap.Controls.Add(barOps);

        Controls.Add(midWrap);
        Controls.Add(upperPanel);
        Controls.Add(bottomBar);
        Controls.Add(topBanner);
    }

    private void BuildQuickRoleButtons(string? group)
    {
        string grp = (group ?? "").Trim().ToLowerInvariant();
        var btns = grp switch
        {
            "beam" or "groundbeam" or "lintel" => new (string, string)[] {
                ("+ Chủ Trên", "TOP"), ("+ Chủ Dưới", "BOT"), ("+ Giá Thành (2M)", "SIDE"),
                ("+ G.Cường Gối", "TOP-GOI"), ("+ G.Cường Nhịp", "BOT-NHIP"),
                ("+ Đai Kín (@1/@2)", "STIR.BEAM"), ("+ Đai Móc C", "TIE.C.BEAM"), ("+ Thanh Tự Do", "CUSTOM")
            },
            "column" => new (string, string)[] {
                ("+ Chủ Nhóm A", "MAIN.A"), ("+ Chủ Nhóm B", "MAIN.B"),
                ("+ Đai Bao Ngoài", "STIR.COL"), ("+ Đai Lồng Trong", "STIR.IN"),
                ("+ Đai Móc C", "TIE.C.COL"), ("+ Thanh Tự Do", "CUSTOM")
            },
            "wall" => new (string, string)[] {
                ("+ Đứng 2 Mặt", "MAIN.V"), ("+ Ngang 2 Mặt", "MAIN.H"),
                ("+ Đai C Lưới", "TIE.C.WALL"), ("+ Thanh Tự Do", "CUSTOM")
            },
            "slab" or "slabongrade" => new (string, string)[] {
                ("+ Dưới X", "BOT.X"), ("+ Dưới Y", "BOT.Y"),
                ("+ Trên X", "TOP.X"), ("+ Trên Y", "TOP.Y"),
                ("+ Mũ Gối X", "HAT.X"), ("+ Mũ Gối Y", "HAT.Y"),
                ("+ Chân Chó", "CHAIR"), ("+ Thanh Tự Do", "CUSTOM")
            },
            "stair" => new (string, string)[] {
                ("+ Thép Nghiêng", "ST-01"), ("+ Mũ Gối", "ST-03"),
                ("+ Phân Bố", "ST-04"), ("+ Thanh Tự Do", "CUSTOM")
            },
            _ => new (string, string)[] {
                ("+ Đáy X", "BOT.X"), ("+ Đáy Y", "BOT.Y"),
                ("+ Trên X", "TOP.X"), ("+ Trên Y", "TOP.Y"),
                ("+ Đai Quây Hông", "AROUND"), ("+ Chân Chó", "CHAIR"), ("+ Thanh Tự Do", "CUSTOM")
            }
        };

        foreach (var (cap, key) in btns)
        {
            var b = new Button
            {
                Text = cap, Width = 100, Height = 30, Margin = new Padding(3),
                BackColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8.2f, FontStyle.Regular)
            };
            b.FlatAppearance.BorderColor = Color.FromArgb(180, 198, 220);
            b.Click += (_, _) =>
            {
                _bars.Add(RebarParamCalculator.CreateQuickRoleRow(_req.Element, _req.Spec, key, _bars.Count + 1));
                RecomputeAll();
                PopulateGrid(_bars.Count - 1);
            };
            _quickFlow.Controls.Add(b);
        }
    }

    private static void AddToolBtn(FlowLayoutPanel p, string text, Action onClick)
    {
        var b = new Button { Text = text, AutoSize = true, Height = 26, Margin = new Padding(3, 1, 3, 1), FlatStyle = FlatStyle.System, Font = new Font("Segoe UI", 8.5f) };
        b.Click += (_, _) => onClick();
        p.Controls.Add(b);
    }

    private void ConfigureGrid()
    {
        _grid.Dock = DockStyle.Fill;
        _grid.AllowUserToAddRows = false;
        _grid.AllowUserToDeleteRows = false;
        _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        _grid.MultiSelect = false;
        _grid.RowHeadersWidth = 28;
        _grid.BackgroundColor = Color.White;
        _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
        _grid.ColumnHeadersHeight = 32;
        _grid.EnableHeadersVisualStyles = false;
        _grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(47, 85, 151);
        _grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
        _grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 8.5f, FontStyle.Bold);

        void AddTxt(string name, string hdr, int w, bool ro = false, Color? bg = null)
        {
            var col = new DataGridViewTextBoxColumn { Name = name, HeaderText = hdr, Width = w, ReadOnly = ro };
            if (bg.HasValue) col.DefaultCellStyle.BackColor = bg.Value;
            if (ro) col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            _grid.Columns.Add(col);
        }
        void AddCombo(string name, string hdr, int w, string[] items)
        {
            var col = new DataGridViewComboBoxColumn { Name = name, HeaderText = hdr, Width = w, FlatStyle = FlatStyle.Flat, DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton };
            col.Items.AddRange(items);
            col.DefaultCellStyle.BackColor = Color.FromArgb(255, 252, 232);
            _grid.Columns.Add(col);
        }

        AddTxt("BarId", "BarID", 76, false, Color.FromArgb(255, 252, 232));
        AddTxt("Role", "Vị trí", 82, false, Color.FromArgb(255, 252, 232));
        AddCombo("Shape", "Dáng", 62, Enumerable.Range(1, 41).Select(i => $"D{i:00}").ToArray());
        AddCombo("Dia", "Ø (mm)", 62, new[] { "6", "8", "10", "12", "14", "16", "18", "20", "22", "25", "28", "32", "36", "40" });
        AddTxt("Part", "Mặt", 44, false, Color.FromArgb(255, 252, 232));
        AddCombo("BarDir", "Phương thanh", 104, new[] { "X", "Y", "H", "X*0.25", "Y*0.25", "X*0.7", "PERIM_YH", "PERIM_XY", "PERIM_XY_INNER", "TIE_Y", "TIE_X", "SLOPED_XH", "CHAIR_H", "CUSTOM" });
        AddCombo("DistDir", "Phương rải", 86, new[] { "X", "Y", "H", "ZONE2_X", "ZONE2_H", "XY_GRID", "XH_GRID", "FIXED" });
        AddTxt("HookL", "Mỏ L", 56, false, Color.FromArgb(255, 252, 232));
        AddTxt("HookR", "Mỏ R", 56, false, Color.FromArgb(255, 252, 232));
        AddTxt("Bv", "Bv", 54, false, Color.FromArgb(255, 252, 232));
        AddTxt("Sp1", "@1", 52, false, Color.FromArgb(255, 252, 232));
        AddTxt("Sp2", "@2", 52, false, Color.FromArgb(255, 252, 232));
        AddTxt("SlOvr", "SL.m", 52, false, Color.FromArgb(255, 245, 210));
        AddTxt("LOvr", "L.m", 62, false, Color.FromArgb(255, 245, 210));
        AddTxt("dPb", "±Pb", 54, false, Color.FromArgb(255, 252, 232));
        AddTxt("dLth", "±Lth", 72, false, Color.FromArgb(255, 252, 232));
        AddTxt("Le1a", "Le.1a", 54, false, Color.White);
        AddTxt("Ri2a", "Ri.2a", 54, false, Color.White);

        Color calcBg = Color.FromArgb(238, 247, 235);
        AddTxt("cLeft", "Left", 56, true, calcBg);
        AddTxt("cMid", "Mid", 66, true, calcBg);
        AddTxt("cRight", "Right", 56, true, calcBg);
        AddTxt("cQty", "SL", 50, true, calcBg);
        AddTxt("cLbar", "Lbar", 66, true, calcBg);
        AddTxt("cKg", "Rebar(kg)", 76, true, Color.FromArgb(226, 239, 218));
        AddTxt("Note", "Ghi chú", 170, false, Color.White);

        _grid.CurrentCellDirtyStateChanged += (_, _) =>
        {
            if (_grid.IsCurrentCellDirty && _grid.CurrentCell is DataGridViewComboBoxCell)
                _grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
        };
        _grid.CellValueChanged += (_, e) => OnGridCellChanged(e.RowIndex, e.ColumnIndex);
        _grid.SelectionChanged += (_, _) => UpdateSelectionVisuals();
        _grid.DataError += (_, e) => { e.Cancel = true; };
    }

    private int SelectedRowIdx() => _grid.CurrentRow?.Index ?? (_bars.Count > 0 ? 0 : -1);

    private void RecomputeAll()
    {
        _calcs.Clear();
        double sumKg = 0;
        foreach (var b in _bars)
        {
            var c = RebarParamCalculator.ComputeRow(_req.Spec, _req.Element, b);
            _calcs.Add(c);
            sumKg += c.WeightKg;
        }
        _lblTotalKg.Text = $"Tổng thép cấu kiện ({_req.Element.Qty:0.##} cái):   {sumKg:N2} kg   ({sumKg / 1000.0:N3} tấn)";
    }

    private void PopulateGrid(int selectIdx = 0)
    {
        _suppressGrid = true;
        _grid.Rows.Clear();
        for (int i = 0; i < _bars.Count; i++)
        {
            var b = _bars[i];
            var c = _calcs[i];
            int idx = _grid.Rows.Add();
            var r = _grid.Rows[idx];
            r.HeaderCell.Value = (i + 1).ToString(CultureInfo.InvariantCulture);
            WriteGridRow(r, b, c);
        }
        if (_grid.Rows.Count > 0)
        {
            int target = Math.Clamp(selectIdx, 0, _grid.Rows.Count - 1);
            _grid.CurrentCell = _grid.Rows[target].Cells[0];
        }
        _suppressGrid = false;
        UpdateSelectionVisuals();
    }

    private static void SetComboSafe(DataGridViewRow r, string colName, string val)
    {
        if (r.Cells[colName] is DataGridViewComboBoxCell cb && !string.IsNullOrEmpty(val) && !cb.Items.Contains(val))
            cb.Items.Add(val);
        r.Cells[colName].Value = val;
    }

    private static void WriteGridRow(DataGridViewRow r, RebarConfigRowDto b, RebarCalcPreview c)
    {
        r.Cells["BarId"].Value = b.BarId;
        r.Cells["Role"].Value = b.Role;
        SetComboSafe(r, "Shape", b.ShapeCode);
        SetComboSafe(r, "Dia", ((int)b.Dia).ToString(CultureInfo.InvariantCulture));
        r.Cells["Part"].Value = b.LayerMult.ToString("0.##", CultureInfo.InvariantCulture);
        SetComboSafe(r, "BarDir", b.BarDir);
        SetComboSafe(r, "DistDir", b.DistDir);
        r.Cells["HookL"].Value = b.HookLExpr;
        r.Cells["HookR"].Value = b.HookRExpr;
        r.Cells["Bv"].Value = b.CoverExpr;
        r.Cells["Sp1"].Value = b.Space1Mm?.ToString("0", CultureInfo.InvariantCulture) ?? "";
        r.Cells["Sp2"].Value = b.Space2Mm?.ToString("0", CultureInfo.InvariantCulture) ?? "";
        r.Cells["SlOvr"].Value = b.OverrideQtyExpr;
        r.Cells["LOvr"].Value = b.OverrideLenExpr;
        r.Cells["dPb"].Value = b.DeltaDistExpr;
        r.Cells["dLth"].Value = b.DeltaLenExpr;
        r.Cells["Le1a"].Value = b.Le1aMm?.ToString("0", CultureInfo.InvariantCulture) ?? "";
        r.Cells["Ri2a"].Value = b.Ri2aMm?.ToString("0", CultureInfo.InvariantCulture) ?? "";
        WriteComputedCells(r, c);
        r.Cells["Note"].Value = b.Note;
    }

    private static void WriteComputedCells(DataGridViewRow r, RebarCalcPreview c)
    {
        r.Cells["cLeft"].Value = c.LeftMm.ToString("0", CultureInfo.InvariantCulture);
        r.Cells["cMid"].Value = c.MidMm.ToString("0", CultureInfo.InvariantCulture);
        r.Cells["cRight"].Value = c.RightMm.ToString("0", CultureInfo.InvariantCulture);
        r.Cells["cQty"].Value = c.QtyBar.ToString("0", CultureInfo.InvariantCulture);
        r.Cells["cLbar"].Value = c.LbarMm.ToString("0", CultureInfo.InvariantCulture);
        r.Cells["cKg"].Value = c.WeightKg.ToString("N2", CultureInfo.InvariantCulture);
    }

    private void OnGridCellChanged(int rowIdx, int colIdx)
    {
        if (_suppressGrid || rowIdx < 0 || rowIdx >= _bars.Count) return;
        var gr = _grid.Rows[rowIdx];
        var b = _bars[rowIdx];
        string Str(string name) => Convert.ToString(gr.Cells[name].Value, CultureInfo.InvariantCulture)?.Trim() ?? "";
        double? NumNull(string name) => double.TryParse(Str(name), NumberStyles.Any, CultureInfo.InvariantCulture, out double v) && v > 0 ? v : null;

        b.BarId = Str("BarId");
        b.Role = Str("Role");
        b.ShapeCode = Str("Shape");
        if (double.TryParse(Str("Dia"), NumberStyles.Any, CultureInfo.InvariantCulture, out double d) && d > 0) b.Dia = d;
        if (double.TryParse(Str("Part"), NumberStyles.Any, CultureInfo.InvariantCulture, out double p) && p > 0) b.LayerMult = p;
        b.BarDir = Str("BarDir");
        b.DistDir = Str("DistDir");
        b.HookLExpr = Str("HookL");
        b.HookRExpr = Str("HookR");
        b.CoverExpr = Str("Bv");
        b.Space1Mm = NumNull("Sp1");
        b.Space2Mm = NumNull("Sp2");
        b.OverrideQtyExpr = Str("SlOvr");
        b.OverrideLenExpr = Str("LOvr");
        b.DeltaDistExpr = Str("dPb");
        b.DeltaLenExpr = Str("dLth");
        b.Le1aMm = NumNull("Le1a");
        b.Ri2aMm = NumNull("Ri2a");
        b.Note = Str("Note");

        _suppressGrid = true;
        RecomputeAll();
        WriteComputedCells(gr, _calcs[rowIdx]);
        _suppressGrid = false;
        UpdateSelectionVisuals();
    }

    private void UpdateSelectionVisuals()
    {
        int idx = SelectedRowIdx();
        if (idx >= 0 && idx < _bars.Count && idx < _calcs.Count)
        {
            var b = _bars[idx];
            var c = _calcs[idx];
            _lblShapeDetail.Text =
                $"Thanh: {b.BarId} ({b.Role})\n" +
                $"Đường kính: Ø{(int)b.Dia} | Dáng: {b.ShapeCode}\n" +
                $"Lbar = {c.LbarMm:0} mm | Lap = {c.LapMm:0} mm\n" +
                $"SL: {c.QtyBar:0} × {(int)b.LayerMult} mặt × {_req.Element.Qty:0.##} CK\n" +
                $"Khối lượng: {c.WeightKg:N2} kg";
        }
        else
        {
            _lblShapeDetail.Text = "";
        }
        _diagramPanel.Invalidate();
        _shapePanel.Invalidate();
    }

    private void DuplicateSelected()
    {
        int idx = SelectedRowIdx();
        if (idx < 0 || idx >= _bars.Count) return;
        var copy = CloneRow(_bars[idx]);
        copy.BarId = copy.BarId + "_2";
        _bars.Insert(idx + 1, copy);
        RecomputeAll();
        PopulateGrid(idx + 1);
    }

    private void MoveSelected(int delta)
    {
        int idx = SelectedRowIdx();
        int target = idx + delta;
        if (idx < 0 || target < 0 || target >= _bars.Count) return;
        (_bars[idx], _bars[target]) = (_bars[target], _bars[idx]);
        RecomputeAll();
        PopulateGrid(target);
    }

    private void DeleteSelected()
    {
        int idx = SelectedRowIdx();
        if (idx < 0 || idx >= _bars.Count) return;
        _bars.RemoveAt(idx);
        RecomputeAll();
        PopulateGrid(Math.Min(idx, _bars.Count - 1));
    }

    private void ReloadPreset()
    {
        _bars.Clear();
        _bars.AddRange(new RebarBbsGenerator().ExpandSingleElementToConfigRows(_req.Element));
        if (_bars.Count == 0)
            _bars.Add(RebarParamCalculator.CreateQuickRoleRow(_req.Element, _req.Spec, DefaultQuickKey(_req.Element.Group), 1));
        RecomputeAll();
        PopulateGrid(0);
    }

    public IReadOnlyList<RebarConfigRowDto> CurrentBars => _bars;

    public void AddQuickRole(string roleKey)
    {
        _bars.Add(RebarParamCalculator.CreateQuickRoleRow(_req.Element, _req.Spec, roleKey, _bars.Count + 1));
        RecomputeAll();
        PopulateGrid(_bars.Count - 1);
    }

    public void AcceptResult() => OnAccept();

    private void OnAccept()
    {
        _grid.EndEdit();
        Result = new EditRebarUiResponse
        {
            Accepted = true,
            ElementId = _req.Element.ElementId,
            Rows = _bars.Select(CloneRow).ToList()
        };
        DialogResult = DialogResult.OK;
        Close();
    }
}
