using System;
using System.Windows.Forms;

namespace WinAppATS
{
    class CopyPaste
    {
        DataGridView dgv;
        public CopyPaste(DataGridView dgv)
        {
            this.dgv = dgv;
        }
        public void CopyToClipboard()
        {
            //Copy to clipboard
            DataObject dataObj = dgv.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }
        public void PasteClipboard(string from)
        {
            try
            {
                string s = Clipboard.GetText();
                string[] lines = s.Split('\n');

                int iRow = dgv.CurrentCell.RowIndex;
                int iCol = dgv.CurrentCell.ColumnIndex;
                DataGridViewCell oCell;

                int iNewRows = iRow + lines.Length;

                //Verify row ultimate containg nothing
                int lenghtRowToCopy = lines[lines.Length - 1] == "" ? iNewRows - 1 : iNewRows;

                //Add rows if is necesary
                if (lenghtRowToCopy > dgv.Rows.Count)
                {
                    dgv.Rows.Add(lenghtRowToCopy - dgv.Rows.Count);
                    if (from.Contains("ventas") || from.Contains("compras"))
                    {
                        for (int i = dgv.Rows.Count - (lines.Length - 2); i < dgv.Rows.Count; i++)
                        {
                            if (from.Contains("ventas"))
                            {
                                for (int j = 6; j < 16; j++)
                                {
                                    oCell = dgv[j, i];
                                    oCell.Value = 0;
                                }
                            }
                            else
                            {
                                for (int j = 11; j < 24; j++)
                                {
                                    oCell = dgv[j, i];
                                    oCell.Value = 0;
                                }
                                oCell = dgv[31, i];
                                oCell.Value = 0;
                            }
                        }
                    }
                }
                //Copy one value in many cells
                if (dgv.SelectedCells.Count > 1)
                {
                    int max = dgv.SelectedCells.Count;
                    s = s.Trim();
                    for (int i = 0; i < max; i++)
                    {
                        oCell = dgv[iCol, iRow - max + i + 1];
                        oCell.Value = s;
                    }
                }
                //Copy values row by row
                else
                {
                    foreach (string line in lines)
                    {
                        if (iRow < dgv.RowCount && line.Length > 0)
                        {
                            string[] sCells = line.Split('\t');
                            for (int i = 0; i < sCells.GetLength(0); ++i)
                            {
                                if (iCol + i < this.dgv.ColumnCount)
                                {
                                    oCell = dgv[iCol + i, iRow];
                                    string notformat = sCells[i].Replace("\r", "");
                                    notformat = notformat.Trim();

                                    if ((from.Contains("compras") && ((iCol + i > 10 && iCol + i < 24) || iCol + i == 31)) ||
                                        (from.Contains("ventas") && (iCol + i > 5 && iCol + i < 16)))
                                    {
                                        notformat = notformat.Replace('.', ',');
                                        notformat = notformat.Replace('-', '0');
                                        if (notformat.Equals(""))
                                        {
                                            notformat = "0";
                                        }
                                        notformat = double.Parse(notformat).ToString();
                                    }
                                    oCell.Value = Convert.ChangeType(notformat, oCell.ValueType);
                                }
                                else
                                {
                                    break;
                                }
                            }
                            iRow++;
                        }
                        else
                        {
                            break;
                        }
                    }
                }
                //    //Clipboard.Clear();
            }
            catch (FormatException)
            {
                MessageBox.Show("¡Error al copiar!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }

        ~CopyPaste() { }
    }
}
