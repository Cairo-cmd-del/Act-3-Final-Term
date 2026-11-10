using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace Bulaong_Sample
{
    public class CrudForm : Form
    {
        private readonly Color BackBlue = Color.FromArgb(155, 178, 216);
        private readonly Color Cream = Color.FromArgb(245, 243, 205);

        // Data is saved here (next to the .exe) so it persists between runs
        private readonly string _file = Path.Combine(Application.StartupPath, "products.xml");

        private DataTable _products;
        private DataGridView grid;
        private TextBox txtId, txtName, txtSupplier, txtSearch;
        private ComboBox cboCategory;
        private NumericUpDown numQty, numPrice;

        public CrudForm()
        {
            Text = "Product Management (CRUD)";
            Size = new Size(1100, 700);
            MinimumSize = new Size(900, 550);
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = BackBlue;

            CreateTable();
            LoadData();
            BuildUI();
            ClearFields();
        }

        // ======================================================= DATA
        private void CreateTable()
        {
            _products = new DataTable("Products");
            _products.Columns.Add("ID", typeof(string));
            _products.Columns.Add("Product Name", typeof(string));
            _products.Columns.Add("Category", typeof(string));
            _products.Columns.Add("Quantity", typeof(int));
            _products.Columns.Add("Price", typeof(decimal));
            _products.Columns.Add("Supplier", typeof(string));
            _products.Columns.Add("Status", typeof(string));
        }

        private void LoadData()
        {
            try
            {
                if (File.Exists(_file))
                {
                    _products.ReadXml(_file);
                    return;
                }
            }
            catch
            {
                _products.Clear(); // file was broken, start with sample data
            }

            AddRow("P001", "Laptop", "Computer", 15, 35000, "TechS");
            AddRow("P002", "Mouse", "Accessory", 30, 500, "TechS");
            AddRow("P003", "Keyboard", "Accessory", 20, 1200, "GadgetP");
            AddRow("P004", "Monitor", "Computer", 5, 6500, "TechS");
            AddRow("P005", "Printer", "Peripheral", 0, 8000, "OfficeP");
            SaveData();
        }

        private void SaveData()
        {
            try
            {
                _products.AcceptChanges();
                _products.WriteXml(_file, XmlWriteMode.WriteSchema);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not save data:\n" + ex.Message, "Save error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static string GetStatus(int qty)
        {
            if (qty <= 0) return "Out of Stock";
            if (qty <= 5) return "Low Stock";
            return "In Stock";
        }

        private void AddRow(string id, string name, string cat, int qty, decimal price, string supplier)
        {
            _products.Rows.Add(id, name, cat, qty, price, supplier, GetStatus(qty));
        }

        private string NextId()
        {
            int max = 0;
            foreach (DataRow r in _products.Rows)
            {
                int n;
                string id = r["ID"].ToString();
                if (id.Length > 1 && int.TryParse(id.Substring(1), out n) && n > max) max = n;
            }
            return "P" + (max + 1).ToString("000");
        }

        // ========================================================= UI
        private void BuildUI()
        {
            // Grid (Fill) must be added first
            grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = Cream,
                BorderStyle = BorderStyle.None,
                GridColor = Color.Gray,
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 36,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                RowTemplate = { Height = 34 }
            };
            grid.ColumnHeadersDefaultCellStyle.BackColor = Cream;
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
            grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10f);
            grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = Cream;
            grid.ColumnHeadersDefaultCellStyle.SelectionForeColor = Color.Black;
            grid.DefaultCellStyle.BackColor = Cream;
            grid.DefaultCellStyle.ForeColor = Color.Black;
            grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(210, 205, 140);
            grid.DefaultCellStyle.SelectionForeColor = Color.Black;
            grid.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            grid.DefaultCellStyle.Font = new Font("Segoe UI", 10f);

            grid.DataBindingComplete += (s, e) =>
            {
                grid.Columns["Price"].DefaultCellStyle.Format = "$#,##0.00";
            };
            grid.CellFormatting += Grid_CellFormatting;
            grid.DataSource = _products;

            // ---- Input fields
            txtId = new TextBox { ReadOnly = true, Dock = DockStyle.Fill, BackColor = Color.Gainsboro };
            txtName = new TextBox { Dock = DockStyle.Fill };
            cboCategory = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDown };
            cboCategory.Items.AddRange(new object[] { "Computer", "Accessory", "Peripheral" });
            numQty = new NumericUpDown { Dock = DockStyle.Fill, Maximum = 1000000, Minimum = 0 };
            numPrice = new NumericUpDown
            {
                Dock = DockStyle.Fill,
                Maximum = 100000000,
                Minimum = 0,
                DecimalPlaces = 2,
                ThousandsSeparator = true
            };
            txtSupplier = new TextBox { Dock = DockStyle.Fill };

            var fields = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 90,
                ColumnCount = 6,
                RowCount = 2,
                Padding = new Padding(5)
            };
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 75));
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.3f));
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 75));
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.3f));
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 75));
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.3f));
            fields.RowStyles.Add(new RowStyle(SizeType.Percent, 50f));
            fields.RowStyles.Add(new RowStyle(SizeType.Percent, 50f));

            fields.Controls.Add(FieldLabel("ID"), 0, 0);
            fields.Controls.Add(txtId, 1, 0);
            fields.Controls.Add(FieldLabel("Name"), 2, 0);
            fields.Controls.Add(txtName, 3, 0);
            fields.Controls.Add(FieldLabel("Category"), 4, 0);
            fields.Controls.Add(cboCategory, 5, 0);
            fields.Controls.Add(FieldLabel("Quantity"), 0, 1);
            fields.Controls.Add(numQty, 1, 1);
            fields.Controls.Add(FieldLabel("Price"), 2, 1);
            fields.Controls.Add(numPrice, 3, 1);
            fields.Controls.Add(FieldLabel("Supplier"), 4, 1);
            fields.Controls.Add(txtSupplier, 5, 1);

            // ---- Buttons + search
            var bar = new Panel { Dock = DockStyle.Top, Height = 50, Padding = new Padding(5, 8, 5, 8) };

            txtSearch = new TextBox { Dock = DockStyle.Right, Width = 200, Font = new Font("Segoe UI", 10f) };
            txtSearch.TextChanged += (s, e) => ApplySearch();
            var lblSearch = new Label { Dock = DockStyle.Right, Width = 60, Text = "Search:", TextAlign = ContentAlignment.MiddleRight };

            var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false };
            buttons.Controls.Add(MakeButton("Add (Create)", Color.ForestGreen, BtnAdd_Click));
            buttons.Controls.Add(MakeButton("Update", Color.SteelBlue, BtnUpdate_Click));
            buttons.Controls.Add(MakeButton("Delete", Color.IndianRed, BtnDelete_Click));
            buttons.Controls.Add(MakeButton("Clear", Color.Gray, (s, e) => ClearFields()));

            bar.Controls.Add(buttons);
            bar.Controls.Add(lblSearch);
            bar.Controls.Add(txtSearch);

            var title = new Label
            {
                Dock = DockStyle.Top,
                Height = 45,
                Text = "Product Management",
                Font = new Font("Segoe UI", 18f, FontStyle.Bold),
                ForeColor = Color.FromArgb(30, 30, 60),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(8, 0, 0, 0)
            };

            Controls.Add(grid);
            Controls.Add(bar);
            Controls.Add(fields);
            Controls.Add(title);

            // Click a row -> show it in the fields (so you can Update / Delete it)
            grid.SelectionChanged += (s, e) => LoadSelectedRow();
        }

        private Label FieldLabel(string text)
        {
            return new Label { Text = text, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
        }

        private Button MakeButton(string text, Color color, EventHandler click)
        {
            var b = new Button
            {
                Text = text,
                Size = new Size(120, 32),
                BackColor = color,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            b.FlatAppearance.BorderSize = 0;
            b.Click += click;
            return b;
        }

        private void Grid_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0 || grid.Columns[e.ColumnIndex].Name != "Status" || e.Value == null) return;
            string s = e.Value.ToString();
            if (s == "Low Stock") e.CellStyle.ForeColor = Color.DarkOrange;
            else if (s == "Out of Stock") e.CellStyle.ForeColor = Color.Red;
            else e.CellStyle.ForeColor = Color.DarkGreen;
        }

        // ============================================== CRUD OPERATIONS
        private void ClearFields()
        {
            txtId.Text = NextId();
            txtName.Clear();
            cboCategory.Text = "";
            numQty.Value = 0;
            numPrice.Value = 0;
            txtSupplier.Clear();
            grid.ClearSelection();
            txtName.Focus();
        }

        private bool ValidateInput()
        {
            if (string.IsNullOrWhiteSpace(txtName.Text))
            {
                MessageBox.Show("Please enter a product name.", "Missing data",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtName.Focus();
                return false;
            }
            if (string.IsNullOrWhiteSpace(cboCategory.Text))
            {
                MessageBox.Show("Please enter a category.", "Missing data",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cboCategory.Focus();
                return false;
            }
            return true;
        }

        // READ: show the selected row in the input fields
        private void LoadSelectedRow()
        {
            if (grid.CurrentRow == null || !grid.CurrentRow.Selected) return;
            var view = grid.CurrentRow.DataBoundItem as DataRowView;
            if (view == null) return;

            DataRow r = view.Row;
            txtId.Text = r["ID"].ToString();
            txtName.Text = r["Product Name"].ToString();
            cboCategory.Text = r["Category"].ToString();
            numQty.Value = Convert.ToDecimal(r["Quantity"]);
            numPrice.Value = Convert.ToDecimal(r["Price"]);
            txtSupplier.Text = r["Supplier"].ToString();
        }

        // CREATE
        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (!ValidateInput()) return;

            AddRow(NextId(), txtName.Text.Trim(), cboCategory.Text.Trim(),
                (int)numQty.Value, numPrice.Value, txtSupplier.Text.Trim());
            SaveData();
            ClearFields();
            MessageBox.Show("Product added.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // UPDATE
        private void BtnUpdate_Click(object sender, EventArgs e)
        {
            DataRow row = FindRowById(txtId.Text);
            if (row == null)
            {
                MessageBox.Show("Click a product in the table first, then press Update.",
                    "No product selected", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (!ValidateInput()) return;

            row["Product Name"] = txtName.Text.Trim();
            row["Category"] = cboCategory.Text.Trim();
            row["Quantity"] = (int)numQty.Value;
            row["Price"] = numPrice.Value;
            row["Supplier"] = txtSupplier.Text.Trim();
            row["Status"] = GetStatus((int)numQty.Value);
            SaveData();
            ClearFields();
            MessageBox.Show("Product updated.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // DELETE
        private void BtnDelete_Click(object sender, EventArgs e)
        {
            DataRow row = FindRowById(txtId.Text);
            if (row == null)
            {
                MessageBox.Show("Click a product in the table first, then press Delete.",
                    "No product selected", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (MessageBox.Show("Delete " + row["Product Name"] + "?", "Confirm delete",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;

            row.Delete();
            SaveData();
            ClearFields();
        }

        private DataRow FindRowById(string id)
        {
            foreach (DataRow r in _products.Rows)
                if (r.RowState != DataRowState.Deleted && r["ID"].ToString() == id) return r;
            return null;
        }

        private void ApplySearch()
        {
            string t = txtSearch.Text.Trim().Replace("'", "''");
            if (t.Length == 0) { _products.DefaultView.RowFilter = ""; return; }
            _products.DefaultView.RowFilter = string.Format(
                "[Product Name] LIKE '%{0}%' OR [Category] LIKE '%{0}%' OR [Supplier] LIKE '%{0}%' OR [ID] LIKE '%{0}%'", t);
        }
    }
}
