using MathNet.Numerics.LinearAlgebra;
using MathNet.Numerics.LinearAlgebra.Double;
using MathNet.Numerics;
using MathNet.Numerics.Distributions;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MathNet.Numerics.Statistics;

namespace Terry_IN_BA_Regression
{
    public partial class Form1 : Form
    {
        public Boolean textBox1Selected = false;
        public Boolean textBox2Selected = false;
        public string[,] arrayX;
        public string[,] arrayY;
        public string[,] arrayXFilter;
        public string[,] arrayYFilter;
        public LinkedList<string> xVariables = new LinkedList<string>();
        public LinkedList<bool> xVariableStates = new LinkedList<bool>();
        public double n;
        public double k;
        public Microsoft.Office.Interop.Excel.Worksheet newWorksheet;
        public int yR = 0;
        public int yC = 0;
        public int xR = 0;
        public int xC = 0;
        Matrix<double> yCap;
        double SSR;
        double yBar;
        double SSE;
        double SST;
        double MSR;
        double MSE;
        double F;
        double P;
        double RSquare;
        double R;
        double RSqAdj;
        Matrix<double> coeff;
        Matrix<double> SE;
        Matrix<double> tStat;
        Matrix<double> pValue;
        Matrix<double> LL;
        Matrix<double> UL;
        Matrix<double> SDx;
        Matrix<double> standardizedCoefficients;
        double SDy; 
        double dfe;
        double tStar;
        Matrix<double> VIFMatrix;
        int originalRows;

        public Form1()
        {
            InitializeComponent();

            textBox3.Enabled = false;
            textBox3.Text = "95";
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            textBox1Selected = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {
        }

        public void doSelect()
        {
            HashSet<String> columns = new HashSet<String>();
            int totalItems = 0;
            string[,] array = { };
            Microsoft.Office.Interop.Excel.Range range = Globals.ThisAddIn.Application.Selection as Microsoft.Office.Interop.Excel.Range;
            string cellnames = null;
            char[] delimiterChars = { '$' };
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
                        cellnames += " [" + words[1] + words[2] + " , " + value + "] ";
                        columns.Add(words[1]);
                        totalItems++;
                    }
                }

                array = new string[(totalItems/columns.Count), columns.Count];
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
                        cellnames += " [" + words[1] + words[2] + " , " + value + "] ";
                        array[r, c] = value;
                        c++;
                        if (c == columns.Count)
                        {
                            c = 0;
                            r++;
                        }
                    }
                }
            }
            if (textBox1Selected == true)
            {

                textBox1.Text = cellnames;
                arrayY = array;
                yR = totalItems / columns.Count;
                yC = columns.Count;

            }
            else if (textBox2Selected == true) 
            {

                textBox2.Text = cellnames;
                arrayX = array;
                xR = totalItems / columns.Count;
                xC = columns.Count;

            }

        }

        private void textBox1_MouseClick(object sender, MouseEventArgs e)
        {
            textBox1Selected = true;
            textBox2Selected = false;
        }

        private void textBox2_MouseClick(object sender, MouseEventArgs e)
        {
            textBox2Selected = true;
            textBox1Selected = false;
        }


        private void button2_Click(object sender, EventArgs e)
        {
            this.Visible = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (checkForInvalidInput())
            {
                return;
            }

            compute(checkBox2.Checked, checkBox1.Checked);
            createOutputSheet();
            clearCache();
            this.Hide();
        }

        private void createOutputSheet()
        {
            newWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets.Add();
            newWorksheet.Select();
            newWorksheet = ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A1", "A1").Value2 = "LINEAR REGRESSION OUTPUT";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A2", "B2").Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A1", "I1").Cells.EntireColumn.ColumnWidth = 20;

            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A3", "A3").Value2 = "Regression Statistics";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A3", "B3").Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A4", "A4").Value2 = "Multiple R";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A5", "A5").Value2 = "R Square";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A6", "A6").Value2 = "Adjusted R Square";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A7", "A7").Value2 = "Standard Error";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A8", "A8").Value2 = "Observations Read";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A9", "A9").Value2 = "Observations Used";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A9", "B9").Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;

            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A10", "A10").Value2 = "ANOVA SUMMARY TABLE";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A10", "F10").Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("B11", "B11").Value2 = "df";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("C11", "C11").Value2 = "SS";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("D11", "D11").Value2 = "MS";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("E11", "E11").Value2 = "F";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("F11", "F11").Value2 = "p-value";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A11", "F11").Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A12", "A12").Value2 = "Regression";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A13", "A13").Value2 = "Residual";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A14", "A14").Value2 = "Total";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A14", "F14").Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;

            if (checkBox4.Checked)
            {
                ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A15", "I15").Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
            } else
            {
                ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A15", "H15").Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
            }

            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("B16", "B16").Value2 = "Coefficients";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("C16", "C16").Value2 = "Standard Error";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("D16", "D16").Value2 = "t Statistic";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("E16", "E16").Value2 = "p-value";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("F16", "F16").Value2 = "Lower " + textBox3.Text + "%";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("G16", "G16").Value2 = "Upper " + textBox3.Text + "%";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("H16", "H16").Value2 = "VIF";

            if (checkBox4.Checked)
            {
                ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("I16", "I16").Value2 = "Standardized Coefficients";
                ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A16", "I16").Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            } else
            {
                ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A16", "H16").Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            }

            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A17", "A17").Value2 = "Intercept";

            if (checkBox2.Checked)
            {
                ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("B17" , "B17").Value2 = "0.000";
            }

            for (int i = 0; i < coeff.RowCount; i++)
            {
                if (checkBox2.Checked)
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("B" + (18 + i), "B" + (18 + i)).Value2 = "" + coeff.At(i, 0);
                }
                else
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("B" + (17 + i), "B" + (17 + i)).Value2 = "" + coeff.At(i, 0);
                }
            }

            for (int i = 0; i < xVariables.Count; i++)
            {               
                ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A" + (18+i), "A" + (18+i)).Value2 = xVariables.ElementAt(i);
            }

            if (checkBox4.Checked)
            {
                if (checkBox2.Checked)
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A" + (17 + coeff.RowCount), "I" + (17 + coeff.RowCount)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                } else
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A" + (16 + coeff.RowCount), "I" + (16 + coeff.RowCount)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                }
            } else
            {
                if (checkBox2.Checked)
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A" + (17 + coeff.RowCount), "H" + (17 + coeff.RowCount)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                } else
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A" + (16 + coeff.RowCount), "H" + (16 + coeff.RowCount)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                }
            }

            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("B12", "B12").Value2 = "" + xVariables.Count;
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("B13", "B13").Value2 = "" + dfe;
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("B14", "B14").Value2 = "" + (n - 1);
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("C12", "C12").Value2 = "" + SSR;
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("C13", "C13").Value2 = "" + SSE;
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("C14", "C14").Value2 = "" + SST;
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("D12", "D12").Value2 = "" + MSR;
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("D13", "D13").Value2 = "" + MSE;
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("E12", "E12").Value2 = "" + F;
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("F12", "F12").Value2 = "" + P;
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("B4", "B4").Value2 = "" + R;
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("B5", "B5").Value2 = "" + RSquare;
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("B6", "B6").Value2 = "" + RSqAdj;
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("B7", "B7").Value2 = "" + Math.Sqrt(MSE);
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("B8", "B8").Value2 = "" + originalRows;
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("B9", "B9").Value2 = "" + n;


            for (int i = 0; i < SE.RowCount; i++)
            {
                if (checkBox2.Checked)
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("C" + (18 + i), "C" + (18 + i)).Value2 = "" + SE.At(i, 0);
                }
                else
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("C" + (17 + i), "C" + (17 + i)).Value2 = "" + SE.At(i, 0);
                }
            }

            for (int i = 0; i < tStat.RowCount; i++)
            {
                if (checkBox2.Checked)
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("D" + (18 + i), "D" + (18 + i)).Value2 = "" + tStat.At(i, 0);
                }
                else
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("D" + (17 + i), "D" + (17 + i)).Value2 = "" + tStat.At(i, 0);
                }
            }

            for (int i = 0; i < pValue.RowCount; i++)
            {
                if (checkBox2.Checked)
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("E" + (18 + i), "E" + (18 + i)).Value2 = "" + pValue.At(i, 0);
                }
                else
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("E" + (17 + i), "E" + (17 + i)).Value2 = "" + pValue.At(i, 0);
                }
            }

            for (int i = 0; i < LL.RowCount; i++)
            {
                if (checkBox2.Checked)
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("F" + (18 + i), "F" + (18 + i)).Value2 = "" + LL.At(i, 0);
                }
                else
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("F" + (17 + i), "F" + (17 + i)).Value2 = "" + LL.At(i, 0);
                }
            }

            for (int i = 0; i < UL.RowCount; i++)
            {
                if (checkBox2.Checked)
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("G" + (18 + i), "G" + (18 + i)).Value2 = "" + UL.At(i, 0);
                }
                else
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("G" + (17 + i), "G" + (17 + i)).Value2 = "" + UL.At(i, 0);
                }
            }

            for (int i = 0; i < xC; i++)
            {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("H" + (18 + i), "H" + (18 + i)).Value2 = "" + VIFMatrix.At(i, i);
            }

            if (checkBox4.Checked)
            {
                for (int i = 0; i < xVariables.Count; i++)
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("I" + (18 + i), "I" + (18 + i)).Value2 = "" + standardizedCoefficients.At(i, 0);
                }
            }

        }

        public void compute(Boolean CiZchecked, Boolean Lchecked)
        {
            double[,] doubleY;
            if (Lchecked)
            {
                doubleY = new double[yR - 1, yC];
                for (int x = 1; x < arrayY.GetLength(0); x++)
                {
                    for (int y = 0; y < arrayY.GetLength(1); y++)
                    {
                        double.TryParse(arrayY[x, y], out doubleY[x - 1, y]);
                    }
                }
            } else
            {
                doubleY = new double[yR, yC];
                for (int x = 0; x < arrayY.GetLength(0); x++)
                {
                    for (int y = 0; y < arrayY.GetLength(1); y++)
                    {
                        double.TryParse(arrayY[x, y], out doubleY[x, y]);
                    }
                }
            }
            Matrix<double> arrayYConverted = DenseMatrix.OfArray(doubleY);
            n = doubleY.Length;
           // MessageBox.Show(arrayYConverted.ToString());


            double[,] doubleX = { };

            if (CiZchecked)
            {
                if (Lchecked)
                {
                    //xVariables = new string[arrayX.GetLength(1)];
                    for (int y = 0; y < arrayX.GetLength(1); y++)
                    {
                        xVariables.AddLast(arrayX[0,y]);
                        xVariableStates.AddLast(true);
                    }
                    doubleX = new double[xR - 1, xC];
                    for (int x = 1; x < arrayX.GetLength(0); x++)
                    {
                        for (int y = 0; y < arrayX.GetLength(1); y++)
                        {
                            double.TryParse(arrayX[x, y], out doubleX[x - 1, y]);
                        }
                    }
                }
                else
                {
                    doubleX = new double[xR, xC];
                    for (int x = 0; x < arrayX.GetLength(0); x++)
                    {
                        for (int y = 0; y < arrayX.GetLength(1); y++)
                        {
                            double.TryParse(arrayX[x, y], out doubleX[x, y]);
                        }
                    }
                }
            }
            else
            {
                if (Lchecked)
                {
                    //xVariables = new string[arrayX.GetLength(1)];
                    for (int y = 0; y < arrayX.GetLength(1); y++)
                    {
                        xVariables.AddLast(arrayX[0, y]);
                        xVariableStates.AddLast(true);
                    }

                    doubleX = new double[xR - 1, xC + 1];
                    for (int x = 0; x < (arrayX.GetLength(0) - 1); x++)
                    {
                        doubleX[x, 0] = 1;
                    }
                    for (int x = 1; x < arrayX.GetLength(0); x++)
                    {
                        for (int y = 0; y < arrayX.GetLength(1); y++)
                        {
                            double.TryParse(arrayX[x, y], out doubleX[x - 1, y + 1]);
                        }
                    }
                }
                else
                {
                    //xVariables = new string[arrayX.GetLength(1)];
                    for (int y = 0; y < arrayX.GetLength(1); y++)
                    {
                        xVariables.AddLast("Variable X" + (y+1));
                        xVariableStates.AddLast(true);
                    }

                    doubleX = new double[xR, xC + 1];
                    for (int x = 0; x < arrayX.GetLength(0); x++)
                    {
                        doubleX[x, 0] = 1;
                    }
                    for (int x = 0; x < arrayX.GetLength(0); x++)
                    {
                        for (int y = 0; y < arrayX.GetLength(1); y++)
                        {
                            double.TryParse(arrayX[x, y], out doubleX[x, y + 1]);
                        }
                    }
                }
            }
            Matrix<double> arrayXConverted = DenseMatrix.OfArray(doubleX);
           

            if (checkBox2.Checked)
            {
                for (int i = xVariableStates.Count - 1; i >= 0; i--)
                {
                    if (!xVariableStates.ElementAt(i))
                    {
                        arrayXConverted = arrayXConverted.RemoveColumn(i);
                        xC--;
                        xVariables.Remove(xVariables.ElementAt(i));
                    }
                }
            } else
            {
                for (int i = xVariableStates.Count - 1; i >= 0; i--)
                {
                    if (!xVariableStates.ElementAt(i))
                    {
                        //MessageBox.Show(arrayXConverted.ToString());
                        arrayXConverted = arrayXConverted.RemoveColumn(i + 1);
                        xC--;
                        xVariables.Remove(xVariables.ElementAt(i));
                    }
                }
            }            

            //MessageBox.Show(arrayXConverted.ToString());
            Matrix<double> coefficients = ((arrayXConverted.Transpose().Multiply(arrayXConverted)).Inverse()).Multiply(arrayXConverted.Transpose()).Multiply(arrayYConverted);
            coeff = coefficients;
            
            k = xVariables.Count;
            dfe = n - xVariables.Count - 1;

            yCap = arrayXConverted.Multiply(coefficients);
            double ySum = 0;
            for(int i = 0; i < arrayYConverted.RowCount; i++)
            {
                ySum += arrayYConverted.At(i, 0);    
            }
            yBar = ySum / arrayYConverted.RowCount;    
            
            for(int i = 0; i < yCap.RowCount; i++)
            {
                SSR += (yCap.At(i, 0) - yBar) * (yCap.At(i, 0) - yBar);    
            }

            for (int i = 0; i < yCap.RowCount; i++)
            {
                SSE += (arrayYConverted.At(i , 0) - yCap.At(i , 0)) * (arrayYConverted.At(i, 0) - yCap.At(i, 0));
            }

            for (int i = 0; i < yCap.RowCount; i++)
            {
                SST += (arrayYConverted.At(i, 0) - yBar) * (arrayYConverted.At(i, 0) - yBar);
            }

            MSR = SSR / k;

            MSE = SSE / (n - k - 1);

            F = MSR / MSE;

            P = ExcelFunctions.FDist(F, (int)k, (int)(n - k - 1));

            RSquare = SSR / SST;

            R = Math.Sqrt(RSquare);

            RSqAdj = (((n - 1) * RSquare) - k) / (n - k - 1);

            SE = MSE * arrayXConverted.Transpose().Multiply(arrayXConverted).Inverse();

            
            for(int i = 0; i < SE.RowCount; i++)
            {
                SE[i,0] = Math.Sqrt(SE.At(i,i));
            }

            double[,] p = new double[coefficients.RowCount,1];

            for (int i = 0; i < xVariables.Count; i++)
            {
                p[i, 0] = 1;
            }

            tStat = DenseMatrix.OfArray(p);
            pValue = DenseMatrix.OfArray(p);
            LL = DenseMatrix.OfArray(p);
            UL = DenseMatrix.OfArray(p);
            standardizedCoefficients = DenseMatrix.OfArray(p);

            p = new double[1, xVariables.Count];

            for (int i = 0; i < xVariables.Count; i++)
            {
                p[0, i] = 1;
            }

            SDx = DenseMatrix.OfArray(p);

            if (checkBox2.Checked)
            {
                for (int i = 0; i < xVariables.Count; i++)
                {
                    SDx[0, i] = MathNet.Numerics.Statistics.Statistics.StandardDeviation(arrayXConverted.Column(i));
                }
            } else
            {
                for (int i = 0; i < xVariables.Count; i++)
                {
                    SDx[0, i] = MathNet.Numerics.Statistics.Statistics.StandardDeviation(arrayXConverted.Column(i + 1));
                }
            }

            SDy = MathNet.Numerics.Statistics.Statistics.StandardDeviation(arrayYConverted.Column(0));

            if (checkBox2.Checked) {
                    for (int i = 0; i < xVariables.Count; i++)
                    {
                        standardizedCoefficients[i, 0] = coefficients.At(i, 0) * (SDx.At(0, i) / SDy);
                        //MessageBox.Show(coefficients.At(i + 1, 0) + " : " + SDx.At(0, i) + " : " + SDy);
                    }
                } else
            {
                for (int i = 0; i < xVariables.Count; i++)
                {
                    standardizedCoefficients[i, 0] = coefficients.At(i + 1, 0) * (SDx.At(0, i) / SDy);
                    //MessageBox.Show(coefficients.At(i + 1, 0) + " : " + SDx.At(0, i) + " : " + SDy);
                }
            }

            for (int i = 0; i < SE.RowCount; i++)
            {
                tStat[i,0] = coefficients.At(i, 0) / SE.At(i, 0);
            }

            for (int i = 0; i < tStat.RowCount; i++)
            {
                if (tStat[i,0] <= 0)
                {
                    pValue[i, 0] = 2 - 2 * ExcelFunctions.TDist(tStat[i,0] , (int)dfe, 1);
                }
                else if (tStat[i,0] > 0) 
                {
                    pValue[i, 0] = 2 - 2 * (1 - ExcelFunctions.TDist(tStat[i, 0], (int)dfe, 1));
                }
            }

            tStar = Math.Abs(StudentT.InvCDF(0d, 1d, (int)dfe, (1 - (double.Parse(textBox3.Text) / 100)) / 2));
            //MessageBox.Show((tStar).ToString());

            for (int i = 0; i < LL.RowCount; i++)
            {
                LL[i, 0] = coeff[i, 0] - tStar * SE[i, 0];
                UL[i, 0] = coeff[i, 0] + tStar * SE[i, 0];
            }

            //MessageBox.Show(arrayXConverted.ToString());

            if (checkBox2.Checked)
            {

                // VIF
                double[,] w = new double[arrayXConverted.RowCount, arrayXConverted.ColumnCount];
                for (int i = 0; i < arrayXConverted.RowCount; i++)
                {
                    for (int j = 0; j < arrayXConverted.ColumnCount; j++)
                    {
                        w[i, j] = 1;
                    }
                }
                Matrix<double> W = DenseMatrix.OfArray(w);
                //MessageBox.Show(W.ToString());

                double[,] s = new double[1, arrayXConverted.ColumnCount];
                for (int i = 0; i < arrayXConverted.ColumnCount; i++)
                {
                    s[0, i] = 1;
                }
                Matrix<double> S = DenseMatrix.OfArray(s);
                //MessageBox.Show(S.ToString());

                double[,] xbar = new double[1, arrayXConverted.ColumnCount];
                for (int i = 0; i < arrayXConverted.ColumnCount; i++)
                {
                    xbar[0, i] = 1;
                }
                Matrix<double> xBar = DenseMatrix.OfArray(xbar);
                //MessageBox.Show(xBar.ToString());

                for (int i = 0; i < arrayXConverted.ColumnCount; i++)
                {
                    double xB = 0;
                    double sum = 0;
                    for (int j = 0; j < arrayXConverted.RowCount; j++)
                    {
                        sum += arrayXConverted.At(j, i);
                    }
                    xB = sum / arrayXConverted.RowCount;
                    xBar[0, i] = xB;

                    double sjj = 0;
                    for (int j = 0; j < arrayXConverted.RowCount; j++)
                    {
                        sjj += (arrayXConverted.At(j, i) - xB) * (arrayXConverted.At(j, i) - xB);
                    }
                    S[0, i] = sjj;
                }
                //MessageBox.Show(xBar.ToString());
                //MessageBox.Show(S.ToString());

                for (int i = 0; i < arrayXConverted.RowCount; i++)
                {
                    for (int j = 0; j < arrayXConverted.ColumnCount; j++)
                    {
                        W[i, j] = (arrayXConverted[i, j] - xBar[0, j]) / (Math.Sqrt(S[0, j]));
                    }
                }
                //MessageBox.Show(W.ToString());

                VIFMatrix = (W.Transpose().Multiply(W)).Inverse();
                //MessageBox.Show(VIFMatrix.ToString());
            }
            else
            {

                // VIF
                double[,] w = new double[arrayXConverted.RowCount, arrayXConverted.ColumnCount - 1];
                for (int i = 0; i < arrayXConverted.RowCount; i++)
                {
                    for (int j = 0; j < arrayXConverted.ColumnCount - 1; j++)
                    {
                        w[i, j] = 1;
                    }
                }
                Matrix<double> W = DenseMatrix.OfArray(w);
                //MessageBox.Show(W.ToString());

                double[,] s = new double[1, arrayXConverted.ColumnCount - 1];
                for (int i = 0; i < arrayXConverted.ColumnCount - 1; i++)
                {
                    s[0, i] = 1;
                }
                Matrix<double> S = DenseMatrix.OfArray(s);
                //MessageBox.Show(S.ToString());

                double[,] xbar = new double[1, arrayXConverted.ColumnCount - 1];
                for (int i = 0; i < arrayXConverted.ColumnCount - 1; i++)
                {
                    xbar[0, i] = 1;
                }
                Matrix<double> xBar = DenseMatrix.OfArray(xbar);
                //MessageBox.Show(xBar.ToString());

                for (int i = 1; i < arrayXConverted.ColumnCount; i++)
                {
                    double xB = 0;
                    double sum = 0;
                    for (int j = 0; j < arrayXConverted.RowCount; j++)
                    {
                        sum += arrayXConverted.At(j, i);
                    }
                    xB = sum / arrayXConverted.RowCount;
                    xBar[0, i - 1] = xB;

                    double sjj = 0;
                    for (int j = 0; j < arrayXConverted.RowCount; j++)
                    {
                        sjj += (arrayXConverted.At(j, i) - xB) * (arrayXConverted.At(j, i) - xB);
                    }
                    S[0, i - 1] = sjj;
                }
                //MessageBox.Show(xBar.ToString());
                //MessageBox.Show(S.ToString());

                for (int i = 0; i < arrayXConverted.RowCount; i++)
                {
                    for (int j = 1; j < arrayXConverted.ColumnCount; j++)
                    {
                        W[i, j - 1] = (arrayXConverted[i, j] - xBar[0, j - 1]) / (Math.Sqrt(S[0, j - 1]));
                    }
                }
                //MessageBox.Show(W.ToString());

                VIFMatrix = (W.Transpose().Multiply(W)).Inverse();
                //MessageBox.Show(VIFMatrix.ToString());
            }

            
        }

        public bool checkForInvalidInput()
        {
            if (yC > 1)
            {
                MessageBox.Show("Dependent variable has to be a single column with one or more rows !", "Error !" , MessageBoxButtons.OK, MessageBoxIcon.Error);
                return true;
            }

            if (yR != xR) 
            {
                MessageBox.Show("The rows of dependent and independent variables must match !", "Error !", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return true;
            }

            if (checkBox1.Checked)
            {
                double number;
                bool isXNumeric = double.TryParse(arrayX[0, 0], out number);
                bool isYNumeric = double.TryParse(arrayY[0, 0], out number);
                if (isXNumeric || isYNumeric)
                {
                    MessageBox.Show("It looks like you expect labels but the first row of input does not look like having labels !", "Error !", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return true;
                }

                for (int i = 1; i < xR; i++)
                {
                    for (int j = 0; j < xC; j++)
                    {
                        if (arrayX[i,j] != "" && !double.TryParse(arrayX[i, j], out number))
                        {
                            MessageBox.Show("It appears that some of the independent variable input data are non-numeric - Specify only numeric input !", "Error !", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return true;
                        }
                    }
                }

                for (int i = 1; i < yR; i++)
                {
                    for (int j = 0; j < yC; j++)
                    {

                        if (!double.TryParse(arrayY[i, j], out number))
                        {
                            MessageBox.Show("It appears that some of the dependent variable input data are non-numeric - Specify only numeric input !", "Error !", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return true;
                        }

                    }

                }

            }
            else
            {
                double number;
                bool isXNumeric = double.TryParse(arrayX[0, 0], out number);
                bool isYNumeric = double.TryParse(arrayY[0, 0], out number);

                if (!isXNumeric || !isYNumeric)
                {
                    MessageBox.Show("It looks like you are not expecting labels but the first row of input does look like having labels !", "Error !", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return true;
                }

                for (int i = 0; i < xR; i++)
                {
                    for (int j = 0; j < xC; j++)
                    {

                        if (arrayX[i, j] != "" && !double.TryParse(arrayX[i, j], out number))
                        {
                            MessageBox.Show("It appears that some of the independent variable input data are non-numeric - Specify only numeric input !", "Error !", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return true;
                        }

                    }
                }

                for (int i = 0; i < yR; i++)
                {
                    for (int j = 0; j < yC; j++)
                    {

                        if (!double.TryParse(arrayY[i, j], out number))
                        {
                            MessageBox.Show("It appears that some of the dependent variable input data are non-numeric - Specify only numeric input !", "Error !", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return true;
                        }

                    }

                }

            }

            double num;

            if (!double.TryParse(textBox3.Text, out num))
            {
                MessageBox.Show("Confidence level does not look like a number - only numbers please !", "Error !", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return true;
            } else
            {
                if (num < 0 || num > 99.9999)
                {
                    MessageBox.Show("Confidence level should be above 0 and less than 100 !", "Error !", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return true;
                }
            }

            adjustRows();

            return false;
        }

        private void adjustRows()
        {

            if (checkBox1.Checked)
            {
                originalRows = yR - 1;
            } 
            else
            {
                originalRows = yR;
            }

            int i, j;
            HashSet<int> removeRowSet = new HashSet<int>();

                for (i = 0; i < xR; i++)
                {
                    for (j = 0; j < xC; j++)
                    {
                        if (arrayX[i, j] == "")
                        {
                            removeRowSet.Add(i);
                        }
                    }
                }
                arrayXFilter = new string[xR - removeRowSet.Count, xC];
                
                int x = 0, y = 0;
                for (i = 0; i < xR; i++)
                {
                    if (!removeRowSet.Contains(i))
                    {
                        y = 0;
                        for (j = 0; j < xC; j++)
                        {
                            arrayXFilter[x, y] = arrayX[i, j];
                            y++;
                        }
                        x++;
                    }
                }
            arrayX = arrayXFilter;

            arrayYFilter = new string[yR - removeRowSet.Count, yC];

            x = 0; y = 0;
            for (i = 0; i < yR; i++)
            {
                if (!removeRowSet.Contains(i))
                {
                    y = 0;
                    for (j = 0; j < yC; j++)
                    {
                        arrayYFilter[x, y] = arrayY[i, j];
                        y++;
                    }
                    x++;
                }
            }

            arrayY = arrayYFilter;

            xR = yR = arrayYFilter.Length;

        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (!textBox3.Enabled && checkBox3.Checked) {
                textBox3.Enabled = true;
            }
            else if (textBox3.Enabled && !checkBox3.Checked)
            {
                textBox3.Enabled = false;
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {


        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void tabPage4_Click(object sender, EventArgs e)
        {

        }

        private void tabPage5_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (arrayX == null || arrayX.Length == 0)
            {
                MessageBox.Show("Please select independent variables first !", "Error !", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            string[] xVariable = new string[xC];

            double value;
            if (!double.TryParse(arrayX[0, 0], out value))
            {
                for (int i = 0; i < xC; i++)
                {
                    xVariable[i] = arrayX[0, i];
                }
            }
            else
            {
                for (int i = 0; i < xC; i++)
                {
                    xVariable[i] = "X" + (i+1);
                }
            }

            Form2 form = new Form2(xVariable);
            form.Visible = true;
            form.TopMost = true;
        }

        public void InjectXVariableStates(LinkedList<bool> xVariableStates)
        {
            this.xVariableStates = xVariableStates;
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {

        }

        public void clearCache()
        {
            textBox1Selected = false;
            textBox2Selected = false;
            arrayX = new string[0,0];
            arrayY = new string[0,0];
            arrayXFilter = new string[0, 0];
            arrayYFilter = new string[0, 0];
            xVariables = new LinkedList<string>();
            xVariableStates = new LinkedList<bool>();
            n = 0.0;
            k = 0.0;
            yR = 0;
            yC = 0;
            xR = 0;
            xC = 0;
            SSR = 0.0;
            yBar = 0.0;
            SSE = 0.0;
            SST = 0.0;
            MSR = 0.0;
            MSE = 0.0;
            F = 0.0;
            P = 0.0;
            RSquare = 0.0;
            R = 0.0;
            RSqAdj = 0.0;
            SDy = 0.0;
            dfe = 0.0;
            tStar = 0.0;
            originalRows = 0;
        }
    }
}