using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;

namespace WinCCTableControl
{
    [DefaultEvent("CellValueChanged")]
    public class CustomTableControl : UserControl
    {
        private DataGridView grid;

        public event EventHandler<CellValueChangedEventArgs> CellValueChanged;

        public CustomTableControl()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            grid = new DataGridView();
            grid.Dock = DockStyle.Fill;
            grid.AllowUserToAddRows = false;
            grid.RowHeadersVisible = false;
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            grid.CellValueChanged += Grid_CellValueChanged;
            grid.CellDoubleClick += Grid_CellDoubleClick;
            this.Controls.Add(grid);
        }

        private void Grid_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
                grid.BeginEdit(true);
        }

        private void Grid_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            var val = grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
            CellValueChanged?.Invoke(this, new CellValueChangedEventArgs(e.RowIndex, e.ColumnIndex, val));
        }

        [Category("Data")]
        [Description("Number of columns in the table.")]
        public int ColumnCount
        {
            get => grid.ColumnCount;
            set
            {
                grid.ColumnCount = value;
                for (int i = 0; i < value; i++)
                    grid.Columns[i].Name = "Col " + (i + 1);
            }
        }

        [Category("Data")]
        [Description("Number of rows in the table.")]
        public int RowCount
        {
            get => grid.RowCount;
            set
            {
                grid.RowCount = value;
            }
        }

        public void SetCell(int row, int col, object value)
        {
            if (row < 0 || row >= grid.RowCount) return;
            if (col < 0 || col >= grid.ColumnCount) return;
            grid.Rows[row].Cells[col].Value = value;
        }

        public object GetCell(int row, int col)
        {
            if (row < 0 || row >= grid.RowCount) return null;
            if (col < 0 || col >= grid.ColumnCount) return null;
            return grid.Rows[row].Cells[col].Value;
        }

        public void SetColumnHeader(int col, string header)
        {
            if (col < 0 || col >= grid.ColumnCount) return;
            grid.Columns[col].HeaderText = header;
        }

        public void ClearAll()
        {
            foreach (DataGridViewRow r in grid.Rows)
                foreach (DataGridViewCell c in r.Cells)
                    c.Value = null;
        }

        public string ExportCsv(char sep = ',')
        {
            var lines = new List<string>();
            var headers = new List<string>();
            foreach (DataGridViewColumn c in grid.Columns) headers.Add(c.HeaderText);
            lines.Add(string.Join(sep.ToString(), headers));
            foreach (DataGridViewRow r in grid.Rows)
            {
                var cells = new List<string>();
                foreach (DataGridViewCell c in r.Cells)
                    cells.Add(c.Value?.ToString() ?? "");
                lines.Add(string.Join(sep.ToString(), cells));
            }
            return string.Join(Environment.NewLine, lines);
        }
    }

    public class CellValueChangedEventArgs : EventArgs
    {
        public int Row { get; }
        public int Column { get; }
        public object Value { get; }

        public CellValueChangedEventArgs(int row, int col, object value)
        {
            Row = row;
            Column = col;
            Value = value;
        }
    }
}
