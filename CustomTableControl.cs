using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WinCCTableControlCS_Aligned
{
    [ComVisible(true)]
    [Guid("12345678-ABCD-1234-ABCD-1234567890AB")]
    [ProgId("WinCCTableControlCS_Aligned.TableControl")]
    public class TableControl : UserControl
    {
        private DataGridView dataGridView;

        public TableControl()
        {
            dataGridView = new DataGridView();
            dataGridView.Dock = DockStyle.Fill;
            dataGridView.ColumnCount = 5;
            dataGridView.RowCount = 10;
            dataGridView.AllowUserToAddRows = false;
            dataGridView.AllowUserToDeleteRows = false;
            dataGridView.ReadOnly = true;

            for (int i = 0; i < 5; i++)
            {
                dataGridView.Columns[i].Name = "Column " + (i + 1);
                dataGridView.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            for (int i = 0; i < 10; i++)
            {
                dataGridView.Rows[i].HeaderCell.Value = "Row " + (i + 1);
                dataGridView.Rows[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            this.Controls.Add(dataGridView);
        }

        public void SetCellValue(int row, int col, string value)
        {
            if (row >= 0 && row < dataGridView.RowCount && col >= 0 && col < dataGridView.ColumnCount)
            {
                dataGridView.Rows[row].Cells[col].Value = value;
            }
        }

        public string GetCellValue(int row, int col)
        {
            if (row >= 0 && row < dataGridView.RowCount && col >= 0 && col < dataGridView.ColumnCount)
            {
                object val = dataGridView.Rows[row].Cells[col].Value;
                return val != null ? val.ToString() : "";
            }
            return "";
        }
    }
}
