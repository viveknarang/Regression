using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;


namespace Terry_IN_BA_Regression
{
    public class View
    {
        public Microsoft.Office.Interop.Excel.Worksheet newWorksheet;
        public OutputModel model = new OutputModel();

        public View(OutputModel model)
        {
            this.model = model;
        }


        public void createOutputOnASeparateSheet()
        {
            newWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets.Add();
            newWorksheet.Select();
            newWorksheet = ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A1", "A1").Value2 = "LINEAR REGRESSION OUTPUT";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A1", "A1").Cells.Font.Bold = true;

            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A2", "A2").Value2 = "Dependent Variable:";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("B2", "B2").Value2 = model.yVariable;

            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A3", "B3").Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A1", "I1").Cells.EntireColumn.ColumnWidth = 20;

            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A4", "A4").Value2 = "Regression Statistics";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A4", "B4").Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A5", "A5").Value2 = "Multiple R";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A6", "A6").Value2 = "R Square";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A7", "A7").Value2 = "Adjusted R Square";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A8", "A8").Value2 = "Standard Error";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A9", "A9").Value2 = "Observations Read";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A10", "A10").Value2 = "Observations Used";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A10", "B10").Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;

            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A12", "A12").Value2 = "ANOVA SUMMARY TABLE";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A12", "F12").Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("B13", "B13").Value2 = "df";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("C13", "C13").Value2 = "SS";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("D13", "D13").Value2 = "MS";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("E13", "E13").Value2 = "F";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("F13", "F13").Value2 = "p-value";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A13", "F13").Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A14", "A14").Value2 = "Regression";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A15", "A15").Value2 = "Residual";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A16", "A16").Value2 = "Total";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A16", "F16").Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;

            if (model.isStandardizedCoefficientsEnabled)
            {
                ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A17", "I17").Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
            }
            else
            {
                ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A17", "H17").Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
            }

    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("B18", "B18").Value2 = "Coefficients";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("C18", "C18").Value2 = "Standard Error";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("D18", "D18").Value2 = "t Statistic";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("E18", "E18").Value2 = "p-value";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("F18", "F18").Value2 = "Lower " + model.confidenceLevel + "%";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("G18", "G18").Value2 = "Upper " + model.confidenceLevel + "%";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("H18", "H18").Value2 = "VIF";

            if (model.isStandardizedCoefficientsEnabled)
            {
                ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("I18", "I18").Value2 = "Standardized Coefficients";
                ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A18", "I18").Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            }
            else
            {
                ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A18", "H18").Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            }

    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A19", "A19").Value2 = "Intercept";

            if (model.noIntercept)
            {
                ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("B19", "B19").Value2 = "0.000";
            }

            for (int i = 0; i < model.betaCoefficients.RowCount; i++)
            {
                if (model.noIntercept)
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("B" + (20 + i), "B" + (20 + i)).Value2 = "" + model.betaCoefficients.At(i, 0);
                }
                else
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("B" + (19 + i), "B" + (19 + i)).Value2 = "" + model.betaCoefficients.At(i, 0);
                }
            }

            for (int i = 0; i < model.xVariables.Count; i++)
            {
                ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A" + (20 + i), "A" + (20 + i)).Value2 = model.xVariables.ElementAt(i);
            }

            if (model.isStandardizedCoefficientsEnabled)
            {
                if (model.noIntercept)
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A" + (19 + model.betaCoefficients.RowCount), "I" + (19 + model.betaCoefficients.RowCount)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                }
                else
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A" + (18 + model.betaCoefficients.RowCount), "I" + (18 + model.betaCoefficients.RowCount)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                }
            }
            else
            {
                if (model.noIntercept)
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A" + (19 + model.betaCoefficients.RowCount), "H" + (19 + model.betaCoefficients.RowCount)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                }
                else
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A" + (18 + model.betaCoefficients.RowCount), "H" + (18 + model.betaCoefficients.RowCount)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                }
            }

    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("B14", "B14").Value2 = "" + model.xVariables.Count;
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("B15", "B15").Value2 = "" + model.dfe;
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("B16", "B16").Value2 = "" + (model.xVariables.Count + model.dfe);
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("C14", "C14").Value2 = "" + model.SSR;
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("C15", "C15").Value2 = "" + model.SSE;
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("C16", "C16").Value2 = "" + model.SST;
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("D14", "D14").Value2 = "" + model.MSR;
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("D15", "D15").Value2 = "" + model.MSE;
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("E14", "E14").Value2 = "" + model.F;
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("F14", "F14").Value2 = "" + model.P;
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("B5", "B5").Value2 = "" + model.R;
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("B6", "B6").Value2 = "" + model.RSquare;
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("B7", "B7").Value2 = "" + model.RSqAdj;
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("B8", "B8").Value2 = "" + Math.Sqrt(model.MSE);
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("B9", "B9").Value2 = "" + model.originalRows;
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("B10", "B10").Value2 = "" + model.n;


            for (int i = 0; i < model.SE.RowCount; i++)
            {
                if (model.noIntercept)
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("C" + (20 + i), "C" + (20 + i)).Value2 = "" + model.SE.At(i, 0);
                }
                else
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("C" + (19 + i), "C" + (19 + i)).Value2 = "" + model.SE.At(i, 0);
                }
            }

            for (int i = 0; i < model.tStat.RowCount; i++)
            {
                if (model.noIntercept)
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("D" + (20 + i), "D" + (20 + i)).Value2 = "" + model.tStat.At(i, 0);
                }
                else
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("D" + (19 + i), "D" + (19 + i)).Value2 = "" + model.tStat.At(i, 0);
                }
            }

            for (int i = 0; i < model.pValue.RowCount; i++)
            {
                if (model.noIntercept)
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("E" + (20 + i), "E" + (20 + i)).Value2 = "" + model.pValue.At(i, 0);
                }
                else
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("E" + (19 + i), "E" + (19 + i)).Value2 = "" + model.pValue.At(i, 0);
                }
            }

            for (int i = 0; i < model.LL.RowCount; i++)
            {
                if (model.noIntercept)
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("F" + (20 + i), "F" + (20 + i)).Value2 = "" + model.LL.At(i, 0);
                }
                else
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("F" + (19 + i), "F" + (19 + i)).Value2 = "" + model.LL.At(i, 0);
                }
            }

            for (int i = 0; i < model.UL.RowCount; i++)
            {
                if (model.noIntercept)
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("G" + (20 + i), "G" + (20 + i)).Value2 = "" + model.UL.At(i, 0);
                }
                else
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("G" + (19 + i), "G" + (19 + i)).Value2 = "" + model.UL.At(i, 0);
                }
            }

            for (int i = 0; i < model.xVariables.Count; i++)
            {
                ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("H" + (20 + i), "H" + (20 + i)).Value2 = "" + model.VIFMatrix.At(i, i);
            }

            if (model.isStandardizedCoefficientsEnabled)
            {
                for (int i = 0; i < model.xVariables.Count; i++)
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("I" + (20 + i), "I" + (20 + i)).Value2 = "" + model.standardizedCoefficients.At(i, 0);
                }
            }

            // ADVANCED TABLE //

            int advancedTableReference = 18 + model.xVariables.Count + 2;
            int advancedTableIndex = 1;

            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A" + advancedTableReference, "A" + advancedTableReference).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A" + (advancedTableReference + 1), "A" + (advancedTableReference + 1)).Value2 = "Obs #";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A" + (advancedTableReference + 1), "A" + (advancedTableReference + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;

            for (int i = 0; i < model.arrayYConverted.RowCount; i++)
            {
                ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A" + (advancedTableReference + 2 + i), "A" + (advancedTableReference + 2 + i)).Value2 = (i + 1);
            }

    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A" + (advancedTableReference + model.arrayYConverted.RowCount + 1), "A" + (advancedTableReference + model.arrayYConverted.RowCount + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;

            if (model.isOriginalEnabledInAdvancedOptions)
            {
                advancedTableIndex++;
                ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + advancedTableReference, Util.IntToLetters(advancedTableIndex) + advancedTableReference).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1)).Value2 = model.yVariable;
                ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;

                for (int i = 0; i < model.arrayYConverted.RowCount; i++)
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i)).Value2 = model.arrayYConverted[i, 0];
                }

                ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + model.arrayYConverted.RowCount + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + model.arrayYConverted.RowCount + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;

                int correction = 1;
                if (model.noIntercept)
                {
                    correction = 0;
                }

                for (int i = correction; i < model.arrayXConverted.ColumnCount; i++)
                {
                    advancedTableIndex++;
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + advancedTableReference, Util.IntToLetters(advancedTableIndex) + advancedTableReference).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1)).Value2 = model.xVariables.ElementAt(i - correction);
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;

                    for (int j = 0; j < model.arrayXConverted.RowCount; j++)
                    {
                        ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + j), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + j)).Value2 = model.arrayXConverted[j, i];
                    }

                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + model.arrayYConverted.RowCount + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + model.arrayYConverted.RowCount + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                }
            }

            if (model.isPredictedEnabledInAdvancedOptions)
            {
                advancedTableIndex++;
                for (int i = 0; i < model.yCap.RowCount; i++)
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + advancedTableReference, Util.IntToLetters(advancedTableIndex) + advancedTableReference).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1)).Value2 = "Predicted";
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i)).Value2 = model.yCap[i, 0];
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + model.arrayYConverted.RowCount + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + model.arrayYConverted.RowCount + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                }

            }


            if (model.isConfidenceLimitsEnabledInAdvancedOptions)
            {
                advancedTableIndex++;
                for (int i = 0; i < model.arrayXConverted.RowCount; i++)
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + advancedTableReference, Util.IntToLetters(advancedTableIndex + 1) + advancedTableReference).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1)).Value2 = model.confidenceLevel + "% Mean LL";
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex + 1) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex + 1) + (advancedTableReference + 1)).Value2 = model.confidenceLevel + "% Mean UL";

                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex + 1) + (advancedTableReference + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                    for (int j = 0; j < model.arrayXConverted.RowCount; j++)
                    {
                        ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + j), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + j)).Value2 = model.LLMean[j, 0];
                        ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex + 1) + (advancedTableReference + 2 + j), Util.IntToLetters(advancedTableIndex + 1) + (advancedTableReference + 2 + j)).Value2 = model.ULMean[j, 0];
                    }
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + model.arrayYConverted.RowCount + 1), Util.IntToLetters(advancedTableIndex + 1) + (advancedTableReference + model.arrayYConverted.RowCount + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                }

            }



            // END ADVANCED TABLE //

        }



    }
}
