using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Terry_IN_BA_Regression
{
    public class Util
    {
        public static Boolean confidenceLevelInBasicChecked = false;
        public static string confidenceLevelDefaultValue = "95";

        public static string IntToLetters(int value)
        {
            string result = string.Empty;
            while (--value >= 0)
            {
                result = (char)('A' + value % 26) + result;
                value /= 26;
            }
            return result;
        }

        public static void showErrorMessage(String message, String heading)
        {
            System.Windows.Forms.MessageBox.Show(message, heading, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
        }

        public static InputModel doSelectInputFromCurrentSheet()
        {
            ThisAddIn.form.updateStatus("Selecting values from the sheet ....");

            InputModel input = new InputModel();

            HashSet<String> columns = new HashSet<String>();
            int totalItems = 0;
            string[,] array = { };
            Microsoft.Office.Interop.Excel.Range range = Globals.ThisAddIn.Application.Selection as Microsoft.Office.Interop.Excel.Range;
            string cellnames = null;
            char[] delimiterChars = { '$' };
            String[] selectedRangeText = { };

            if (range != null)
            {
                for (int areaIndex = 1; areaIndex <= range.Areas.Count; areaIndex++)
                {
                    if (areaIndex > 1000)
                    {
                        break;
                    }
                    for (int cellIndex = 1; cellIndex <= range.Areas[areaIndex].Cells.Count; cellIndex++)
                    {
                        if (cellIndex > 1000)
                        {
                            break;
                        }
                        string address = range.Areas[areaIndex].Cells[cellIndex].Address;
                        string value = "" + range.Areas[areaIndex].Cells[cellIndex].Value;
                        string[] words = address.Split(delimiterChars);
                        cellnames += "[" + words[1] + words[2] + " , " + value + "] ";
                        columns.Add(words[1]);
                        totalItems++;
                    }
                }
                array = new string[(totalItems / columns.Count), columns.Count];
                cellnames = null;
                int r = 0, c = 0;
                for (int areaIndex = 1; areaIndex <= range.Areas.Count; areaIndex++)
                {
                    if (areaIndex > 1000)
                    {
                        break;
                    }
                    for (int cellIndex = 1; cellIndex <= range.Areas[areaIndex].Cells.Count; cellIndex++)
                    {
                        if (cellIndex > 1000)
                        {
                            break;
                        }
                        string address = range.Areas[areaIndex].Cells[cellIndex].Address;
                        string value = "" + range.Areas[areaIndex].Cells[cellIndex].Value;
                        string[] words = address.Split(delimiterChars);
                        cellnames += "[" + words[1] + words[2] + "," + value + "] ";
                        array[r, c] = value;
                        c++;
                        if (c == columns.Count)
                        {
                            c = 0;
                            r++;
                        }
                    }
                }

                selectedRangeText = System.Text.RegularExpressions.Regex.Split(cellnames," ");
            }
            input.array = array;
            input.cellnames = selectedRangeText[0] + " : " + selectedRangeText[selectedRangeText.Length-2];
            input.columns = columns;
            input.totalItems = totalItems;


            ThisAddIn.form.updateStatus("Selecting values from the sheet .... [OK]");
            return input;
        }
    }
}
