using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace WinCCTableControlCS_StringAlignment
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
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

        [ComVisible(true)]
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

        [ComVisible(true)]
        public int RowCount
        {
            get => grid.RowCount;
            set => grid.RowCount = value;
        }

        private int[] columnWidths = new int[5]; // Adjust size as needed

        [ComVisible(true)]
        [Category("Layout")]
        [Description("Width of Column 0")]
        public int Column0Width
        {
            get => columnWidths[0];
            set
            {
                columnWidths[0] = value;
                if (grid.ColumnCount > 0)
                    grid.Columns[0].Width = value;
            }
        }
        
        [ComVisible(true)]
        [Category("Layout")]
        [Description("Width of Column 1")]
        public int Column1Width
        {
            get => columnWidths[1];
            set
            {
                columnWidths[1] = value;
                if (grid.ColumnCount > 1)
                    grid.Columns[1].Width = value;
            }
        }
        
        [ComVisible(true)]
        [Category("Layout")]
        [Description("Width of Column 2")]
        public int Column2Width
        {
            get => columnWidths[2];
            set
            {
                columnWidths[2] = value;
                if (grid.ColumnCount > 2)
                    grid.Columns[2].Width = value;
            }
        }
        
        [ComVisible(true)]
        [Category("Layout")]
        [Description("Width of Column 3")]
        public int Column3Width
        {
            get => columnWidths[3];
            set
            {
                columnWidths[3] = value;
                if (grid.ColumnCount > 3)
                    grid.Columns[3].Width = value;
            }
        }
        
        [ComVisible(true)]
        [Category("Layout")]
        [Description("Width of Column 4")]
        public int Column4Width
        {
            get => columnWidths[4];
            set
            {
                columnWidths[4] = value;
                if (grid.ColumnCount > 4)
                    grid.Columns[4].Width = value;
            }
        }
        

        [ComVisible(true)]
        public void SetCell(int row, int col, object value)
        {
            if (row < 0 || row >= grid.RowCount) return;
            if (col < 0 || col >= grid.ColumnCount) return;
            grid.Rows[row].Cells[col].Value = value;
        }

        [ComVisible(true)]
        public object GetCell(int row, int col)
        {
            if (row < 0 || row >= grid.RowCount) return null;
            if (col < 0 || col >= grid.ColumnCount) return null;
            return grid.Rows[row].Cells[col].Value;
        }

        [ComVisible(true)]
        public void SetColumnHeader(int col, string header)
        {
            if (col < 0 || col >= grid.ColumnCount) return;
            grid.Columns[col].HeaderText = header;
        }

        [ComVisible(true)]
        public void ClearAll()
        {
            foreach (DataGridViewRow r in grid.Rows)
                foreach (DataGridViewCell c in r.Cells)
                    c.Value = null;
        }

        [ComVisible(true)]
        public void SetColumnAlignment(int colIndex, string alignment)
        {
            if (colIndex < 0 || colIndex >= grid.ColumnCount) return;

            if (Enum.TryParse(alignment, out DataGridViewContentAlignment align))
            {
                grid.Columns[colIndex].DefaultCellStyle.Alignment = align;
                columnAlignments[colIndex] = align;
            }
        }

        [ComVisible(true)]
        public string GetColumnAlignment(int colIndex)
        {
            if (colIndex < 0 || colIndex >= grid.ColumnCount) return grid.DefaultCellStyle.Alignment.ToString();
            return columnAlignments.ContainsKey(colIndex) ? columnAlignments[colIndex].ToString() : grid.Columns[colIndex].DefaultCellStyle.Alignment.ToString();
        }
    }

    [ComVisible(true)]
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
