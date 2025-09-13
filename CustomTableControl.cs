using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;

namespace WinCCTableControl
{
    [DefaultEvent("CellValueChanged")]
    public class CustomTableControl : UserControl
    {
        private DataGridView grid;
        private Dictionary<int, DataGridViewContentAlignment> columnAlignments = new Dictionary<int, DataGridViewContentAlignment>();

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
            grid.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
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

        [Category("Appearance")]
        [Description("Set text alignment for a specific column (by index).")]
        public void SetColumnAlignment(int colIndex, DataGridViewContentAlignment alignment)
        {
            if (colIndex < 0 || colIndex >= grid.ColumnCount) return;
            grid.Columns[colIndex].DefaultCellStyle.Alignment = alignment;
            columnAlignments[colIndex] = alignment;
        }

        public DataGridViewContentAlignment GetColumnAlignment(int colIndex)
        {
            if (colIndex < 0 || colIndex >= grid.ColumnCount) return grid.DefaultCellStyle.Alignment;
            return columnAlignments.ContainsKey(colIndex) ? columnAlignments[colIndex] : grid.Columns[colIndex].DefaultCellStyle.Alignment;
        }

        [Category("Appearance")]
        [Description("Comma-separated text alignment for each column. Example: MiddleCenter,MiddleLeft,MiddleRight")]
        public string ColumnTextAlignments
        {
            get
            {
                return string.Join(",", grid.Columns.Cast<DataGridViewColumn>().Select(c => c.DefaultCellStyle.Alignment.ToString()));
            }
            set
            {
                if (string.IsNullOrWhiteSpace(value)) return;
                var alignments = value.Split(',');
                for (int i = 0; i < Math.Min(alignments.Length, grid.ColumnCount); i++)
                {
                    if (Enum.TryParse(alignments[i].Trim(), out DataGridViewContentAlignment align))
                    {
                        grid.Columns[i].DefaultCellStyle.Alignment = align;
                    }
                }
            }
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
