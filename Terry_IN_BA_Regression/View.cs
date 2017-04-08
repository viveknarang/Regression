using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace Terry_IN_BA_Regression
{
    public class View
    {
        public Microsoft.Office.Interop.Excel.Worksheet newWorksheet;

        public OutputModel model = new OutputModel();
        public InputModel input = new InputModel();

        public View(OutputModel model, InputModel input)
        {
            this.model = model;
            this.input = input;
            this.drawPlots();
        }

        public void drawPlots()
        {
            Microsoft.Office.Tools.Excel.Workbook workbook = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook);
            Sheets charts = workbook.Charts;            
            SeriesCollection seriesCollectionX = null;
            Series seriesX = null;

            Microsoft.Office.Interop.Excel.ChartObjects ChartObjects = (Microsoft.Office.Interop.Excel.ChartObjects)workbook.Sheets.Add().ChartObjects();

            int i = 0;

            int x1 = 100, x3 = 600;

            int globalCounter = 0;

            if (model.isScatterPlotCheckedInPAndGSection)
            {
                for (i = 0; i < model.xVariables.Count; i++)
                {
                    Microsoft.Office.Interop.Excel.ChartObject chartObject;

                    if ((i + 1) % 2 == 0)
                    {
                        chartObject = ChartObjects.Add(x3, 20 + (200 * (globalCounter - ((globalCounter) % 2))), 400, 350);
                    }
                    else
                    {
                        chartObject = ChartObjects.Add(x1, 20 + (200 * globalCounter), 400, 350);
                    }

                    globalCounter++;

                    Chart chart = chartObject.Chart;
                    chart.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlXYScatter;
                    chart.HasTitle = true;
                    chart.ChartTitle.Text = "Scatterplot of " + model.yVariable + " by " + model.xVariables.ElementAt(i);
                    chart.HasLegend = false;
                    seriesCollectionX = (SeriesCollection)chart.SeriesCollection();
                    seriesX = seriesCollectionX.NewSeries();
                    seriesX.Values = model.arrayYConverted.ToArray();
                    seriesX.XValues = model.arrayXConverted.Column(i + 1).ToArray();
                    seriesX.Name = model.xVariables.ElementAt(i);
                    seriesX.MarkerStyle = XlMarkerStyle.xlMarkerStyleCircle;
                    chart.WallsAndGridlines2D = false;
                    Axis axis = (Axis)chart.Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary);
                    axis.HasTitle = true;
                    axis.AxisTitle.Text = model.yVariable;
                    axis.HasMajorGridlines = false;
                    axis.HasMinorGridlines = false;
                    Axis axis2 = (Axis)chart.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
                    axis2.HasTitle = true;
                    axis2.AxisTitle.Text = model.xVariables.ElementAt(i);
                    axis2.HasMajorGridlines = false;
                    axis2.HasMinorGridlines = false;
                    axis2.MinimumScale = model.arrayXConverted.Column(i + 1).Min();
                    axis.MinimumScale = model.arrayYConverted.Column(0).Min();
                    axis.CrossesAt = model.arrayYConverted.Column(0).Min();
                    axis2.CrossesAt = model.arrayXConverted.Column(i + 1).Min();
                }
            }

            if (model.isResidualsByPredictedCheckedInPAndGSection)
            {
                Microsoft.Office.Interop.Excel.ChartObject chartObject1;

                if ((i + 1) % 2 == 0)
                {
                    chartObject1 = ChartObjects.Add(x3, 20 + (200 * (globalCounter - ((globalCounter) % 2))), 400, 350);
                }
                else
                {
                    chartObject1 = ChartObjects.Add(x1, 20 + (200 * globalCounter), 400, 350);
                }

                globalCounter++;

                Chart chart1 = chartObject1.Chart;
                chart1.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlXYScatter;
                chart1.HasTitle = true;
                chart1.ChartTitle.Text = "Residual Plot By Predicted Values of " + model.yVariable;
                chart1.HasLegend = false;
                seriesCollectionX = (SeriesCollection)chart1.SeriesCollection();
                seriesX = seriesCollectionX.NewSeries();
                seriesX.Values = model.residuals.ToArray();
                seriesX.XValues = model.yCap.ToArray();
                seriesX.Name = "Predicted";
                seriesX.MarkerStyle = XlMarkerStyle.xlMarkerStyleCircle;
                chart1.WallsAndGridlines2D = false;
                Axis axis1 = (Axis)chart1.Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary);
                axis1.HasTitle = true;
                axis1.AxisTitle.Text = "Residuals";
                axis1.HasMajorGridlines = false;
                axis1.HasMinorGridlines = false;
                Axis axis21 = (Axis)chart1.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
                axis21.HasTitle = true;
                axis21.AxisTitle.Text = "Predicted Value of " + model.yVariable; ;
                axis21.HasMajorGridlines = false;
                axis21.HasMinorGridlines = false;
                axis21.MinimumScale = model.yCap.Column(0).Min();
                axis1.MinimumScale = model.residuals.Column(0).Min();
                axis1.CrossesAt = model.residuals.Column(0).Min();
                axis21.CrossesAt = model.yCap.Column(0).Min();
            }

            if (model.isStandardizedResidualsByPredictedCheckedInPAndGSection)
            {
                i++;
                Microsoft.Office.Interop.Excel.ChartObject chartObject1;

                if ((i + 1) % 2 == 0)
                {
                    chartObject1 = ChartObjects.Add(x3, 20 + (200 * (globalCounter - ((globalCounter) % 2))), 400, 350);
                }
                else
                {
                    chartObject1 = ChartObjects.Add(x1, 20 + (200 * globalCounter), 400, 350);
                }

                globalCounter++;

                Chart chart1 = chartObject1.Chart;
                chart1.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlXYScatter;
                chart1.HasTitle = true;
                chart1.ChartTitle.Text = "Std. Residual Plot By Predicted Values of " + model.yVariable;
                chart1.HasLegend = false;
                seriesCollectionX = (SeriesCollection)chart1.SeriesCollection();
                seriesX = seriesCollectionX.NewSeries();
                seriesX.Values = model.standardizedResiduals.ToArray();
                seriesX.XValues = model.yCap.ToArray();
                seriesX.Name = "Predicted";
                seriesX.MarkerStyle = XlMarkerStyle.xlMarkerStyleCircle;
                chart1.WallsAndGridlines2D = false;
                Axis axis1 = (Axis)chart1.Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary);
                axis1.HasTitle = true;
                axis1.AxisTitle.Text = "Standardized Residuals";
                axis1.HasMajorGridlines = false;
                axis1.HasMinorGridlines = false;
                Axis axis21 = (Axis)chart1.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
                axis21.HasTitle = true;
                axis21.AxisTitle.Text = "Predicted Value of " + model.yVariable; ;
                axis21.HasMajorGridlines = false;
                axis21.HasMinorGridlines = false;
                axis21.MinimumScale = model.yCap.Column(0).Min();
                axis1.MinimumScale = model.standardizedResiduals.Column(0).Min();
                axis1.CrossesAt = model.standardizedResiduals.Column(0).Min();
                axis21.CrossesAt = model.yCap.Column(0).Min();
            }

            if (model.isResidualsByXVariablesCheckedInPAndGSection)
            {
                i++;
                for (i = 0 ; i < model.xVariables.Count; i++)
                {
                    Microsoft.Office.Interop.Excel.ChartObject chartObject;

                    if ((i + 1) % 2 == 0)
                    {
                        chartObject = ChartObjects.Add(x3, 20 + (200 * (globalCounter - ((globalCounter) % 2))), 400, 350);
                    }
                    else
                    {
                        chartObject = ChartObjects.Add(x1, 20 + (200 * globalCounter), 400, 350);
                    }

                    globalCounter++;

                    Chart chart = chartObject.Chart;
                    chart.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlXYScatter;
                    chart.HasTitle = true;
                    chart.ChartTitle.Text = "Residual Plot by " + model.xVariables.ElementAt(i);
                    chart.HasLegend = false;
                    seriesCollectionX = (SeriesCollection)chart.SeriesCollection();
                    seriesX = seriesCollectionX.NewSeries();
                    seriesX.Values = model.residuals.ToArray();
                    seriesX.XValues = model.arrayXConverted.Column(i + 1).ToArray();
                    seriesX.Name = model.xVariables.ElementAt(i);
                    seriesX.MarkerStyle = XlMarkerStyle.xlMarkerStyleCircle;
                    chart.WallsAndGridlines2D = false;
                    Axis axis = (Axis)chart.Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary);
                    axis.HasTitle = true;
                    axis.AxisTitle.Text = "Residuals";
                    axis.HasMajorGridlines = false;
                    axis.HasMinorGridlines = false;
                    Axis axis2 = (Axis)chart.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
                    axis2.HasTitle = true;
                    axis2.AxisTitle.Text = model.xVariables.ElementAt(i);
                    axis2.HasMajorGridlines = false;
                    axis2.HasMinorGridlines = false;
                    axis2.MinimumScale = model.arrayXConverted.Column(i + 1).Min();
                    axis.MinimumScale = model.residuals.Column(0).Min();
                    axis.CrossesAt = model.residuals.Column(0).Min();
                    axis2.CrossesAt = model.arrayXConverted.Column(i + 1).Min();
                }
            }

            if (model.isStandardizedResidualsByXVariablesCheckedInPAndGSection)
            {
                i++;
                for (i = 0; i < model.xVariables.Count; i++)
                {
                    Microsoft.Office.Interop.Excel.ChartObject chartObject;

                    if ((i + 1) % 2 == 0)
                    {
                        chartObject = ChartObjects.Add(x3, 20 + (200 * (globalCounter - ((globalCounter) % 2))), 400, 350);
                    }
                    else
                    {
                        chartObject = ChartObjects.Add(x1, 20 + (200 * globalCounter), 400, 350);
                    }

                    globalCounter++;

                    Chart chart = chartObject.Chart;
                    chart.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlXYScatter;
                    chart.HasTitle = true;
                    chart.ChartTitle.Text = "Residual Plot by " + model.xVariables.ElementAt(i);
                    chart.HasLegend = false;
                    seriesCollectionX = (SeriesCollection)chart.SeriesCollection();
                    seriesX = seriesCollectionX.NewSeries();
                    seriesX.Values = model.standardizedResiduals.ToArray();
                    seriesX.XValues = model.arrayXConverted.Column(i + 1).ToArray();
                    seriesX.Name = model.xVariables.ElementAt(i);
                    seriesX.MarkerStyle = XlMarkerStyle.xlMarkerStyleCircle;
                    chart.WallsAndGridlines2D = false;
                    Axis axis = (Axis)chart.Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary);
                    axis.HasTitle = true;
                    axis.AxisTitle.Text = "Standardized Residuals";
                    axis.HasMajorGridlines = false;
                    axis.HasMinorGridlines = false;
                    Axis axis2 = (Axis)chart.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
                    axis2.HasTitle = true;
                    axis2.AxisTitle.Text = model.xVariables.ElementAt(i);
                    axis2.HasMajorGridlines = false;
                    axis2.HasMinorGridlines = false;
                    axis2.MinimumScale = model.arrayXConverted.Column(i + 1).Min();
                    axis.MinimumScale = model.standardizedResiduals.Column(0).Min();
                    axis.CrossesAt = model.standardizedResiduals.Column(0).Min();
                    axis2.CrossesAt = model.arrayXConverted.Column(i + 1).Min();
                }
            }

            if (model.isResidualsCheckedInPAndGSection)
            {
                i++;
                Microsoft.Office.Interop.Excel.ChartObject chartObject1;

                if ((i + 1) % 2 == 0)
                {
                    chartObject1 = ChartObjects.Add(x3, 20 + (200 * (globalCounter - ((globalCounter) % 2))), 400, 350);
                }
                else
                {
                    chartObject1 = ChartObjects.Add(x1, 20 + (200 * globalCounter), 400, 350);
                }

                globalCounter++;

                Chart chart1 = chartObject1.Chart;
                chart1.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlXYScatter;
                chart1.HasTitle = true;
                chart1.ChartTitle.Text = "Normal Plot of Model Residuals";
                chart1.HasLegend = false;
                seriesCollectionX = (SeriesCollection)chart1.SeriesCollection();
                seriesX = seriesCollectionX.NewSeries();
                seriesX.Values = model.standardNormalQuantileForResiduals.ToArray();
                seriesX.XValues = model.residuals.ToArray();
                //seriesX.Name = "Residuals";
                seriesX.MarkerStyle = XlMarkerStyle.xlMarkerStyleCircle;
                chart1.WallsAndGridlines2D = false;
                Axis axis1 = (Axis)chart1.Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary);
                axis1.HasTitle = true;
                axis1.AxisTitle.Text = "Theoretical Normal Scores";
                axis1.HasMajorGridlines = false;
                axis1.HasMinorGridlines = false;
                Axis axis21 = (Axis)chart1.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
                axis21.HasTitle = true;
                axis21.AxisTitle.Text = "Residuals";
                axis21.HasMajorGridlines = false;
                axis21.HasMinorGridlines = false;
                axis21.MinimumScale = model.residuals.Column(0).Min();
                axis1.MinimumScale = model.standardNormalQuantileForResiduals.Column(0).Min();
                axis1.CrossesAt = model.standardNormalQuantileForResiduals.Column(0).Min();
                axis21.CrossesAt = model.residuals.Column(0).Min();

            }

            if (model.isStandardizedResidualsByPredictedCheckedInPAndGSection)
            {
                i++;
                Microsoft.Office.Interop.Excel.ChartObject chartObject1;

                if ((i + 1) % 2 == 0)
                {
                    chartObject1 = ChartObjects.Add(x3, 20 + (200 * (globalCounter - ((globalCounter) % 2))), 400, 350);
                }
                else
                {
                    chartObject1 = ChartObjects.Add(x1, 20 + (200 * globalCounter), 400, 350);
                }

                globalCounter++;

                Chart chart1 = chartObject1.Chart;
                chart1.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlXYScatter;
                chart1.HasTitle = true;
                chart1.ChartTitle.Text = "Normal Plot of Standardized Residuals";
                chart1.HasLegend = false;
                seriesCollectionX = (SeriesCollection)chart1.SeriesCollection();
                seriesX = seriesCollectionX.NewSeries();
                seriesX.Values = model.standardNormalQuantileForStandardizedResiduals.ToArray();
                seriesX.XValues = model.standardizedResiduals.ToArray();
                //seriesX.Name = "Residuals";
                seriesX.MarkerStyle = XlMarkerStyle.xlMarkerStyleCircle;
                chart1.WallsAndGridlines2D = false;
                Axis axis1 = (Axis)chart1.Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary);
                axis1.HasTitle = true;
                axis1.AxisTitle.Text = "Theoretical Normal Scores";
                axis1.HasMajorGridlines = false;
                axis1.HasMinorGridlines = false;
                Axis axis21 = (Axis)chart1.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
                axis21.HasTitle = true;
                axis21.AxisTitle.Text = "Standardized Residuals";
                axis21.HasMajorGridlines = false;
                axis21.HasMinorGridlines = false;
                axis21.MinimumScale = model.standardizedResiduals.Column(0).Min();
                axis1.MinimumScale = model.standardNormalQuantileForStandardizedResiduals.Column(0).Min();
                axis1.CrossesAt = model.standardNormalQuantileForStandardizedResiduals.Column(0).Min();
                axis21.CrossesAt = model.standardizedResiduals.Column(0).Min();

            }

            if (model.isYVariableCheckedInPAndGSection)
            {
                i++;
                Microsoft.Office.Interop.Excel.ChartObject chartObject1;

                if ((i + 1) % 2 == 0)
                {
                    chartObject1 = ChartObjects.Add(x3, 20 + (200 * (globalCounter - ((globalCounter) % 2))), 400, 350);
                }
                else
                {
                    chartObject1 = ChartObjects.Add(x1, 20 + (200 * globalCounter), 400, 350);
                }

                globalCounter++;

                Chart chart1 = chartObject1.Chart;
                chart1.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlXYScatter;
                chart1.HasTitle = true;
                chart1.ChartTitle.Text = "Normal Plot of Response";
                chart1.HasLegend = false;
                seriesCollectionX = (SeriesCollection)chart1.SeriesCollection();
                seriesX = seriesCollectionX.NewSeries();
                seriesX.Values = model.standardNormalQuantileForStandardizedResiduals.ToArray();
                seriesX.XValues = model.arrayYConverted.ToArray();
                //seriesX.Name = "Residuals";
                seriesX.MarkerStyle = XlMarkerStyle.xlMarkerStyleCircle;
                chart1.WallsAndGridlines2D = false;
                Axis axis1 = (Axis)chart1.Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary);
                axis1.HasTitle = true;
                axis1.AxisTitle.Text = "Theoretical Normal Scores";
                axis1.HasMajorGridlines = false;
                axis1.HasMinorGridlines = false;
                Axis axis21 = (Axis)chart1.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
                axis21.HasTitle = true;
                axis21.AxisTitle.Text = "" + model.yVariable;
                axis21.HasMajorGridlines = false;
                axis21.HasMinorGridlines = false;
                axis21.MinimumScale = model.arrayYConverted.Column(0).Min();
                axis1.MinimumScale = model.standardNormalQuantileForStandardizedResiduals.Column(0).Min();
                axis1.CrossesAt = model.standardNormalQuantileForStandardizedResiduals.Column(0).Min();
                axis21.CrossesAt = model.arrayYConverted.Column(0).Min();

            }

            if (model.isLeverageCheckedInPAndGSection)
            {
                i++;
                Microsoft.Office.Interop.Excel.ChartObject chartObject1;

                if ((i + 1) % 2 == 0)
                {
                    chartObject1 = ChartObjects.Add(x3, 20 + (200 * (globalCounter - ((globalCounter) % 2))), 400, 350);
                }
                else
                {
                    chartObject1 = ChartObjects.Add(x1, 20 + (200 * globalCounter), 400, 350);
                }

                globalCounter++;

                Chart chart1 = chartObject1.Chart;
                chart1.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlXYScatterLines;
                chart1.HasTitle = true;
                chart1.ChartTitle.Text = "Leverage Chart";
                chart1.HasLegend = false;
                seriesCollectionX = (SeriesCollection)chart1.SeriesCollection();
                seriesX = seriesCollectionX.NewSeries();
                seriesX.Values = model.Leverage.ToArray();
                seriesX.XValues = model.observationNumber.ToArray();
                //seriesX.Name = "Residuals";
                seriesX.MarkerStyle = XlMarkerStyle.xlMarkerStyleCircle;
                chart1.WallsAndGridlines2D = false;
                Axis axis1 = (Axis)chart1.Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary);
                axis1.HasTitle = true;
                axis1.AxisTitle.Text = "Leverage";
                axis1.HasMajorGridlines = false;
                axis1.HasMinorGridlines = false;
                Axis axis21 = (Axis)chart1.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
                axis21.HasTitle = true;
                axis21.AxisTitle.Text = "Observation Number";
                axis21.HasMajorGridlines = false;
                axis21.HasMinorGridlines = false;
                axis21.MinimumScale = model.observationNumber.Column(0).Min();
                axis1.MinimumScale = model.Leverage.Column(0).Min();
                axis1.CrossesAt = model.Leverage.Column(0).Min();
                axis21.CrossesAt = model.observationNumber.Column(0).Min();
            }

            if (model.isCooksDCheckedInPAndGSection)
            {
                i++;
                Microsoft.Office.Interop.Excel.ChartObject chartObject1;

                if ((i + 1) % 2 == 0)
                {
                    chartObject1 = ChartObjects.Add(x3, 20 + (200 * (globalCounter - ((globalCounter) % 2))), 400, 350);
                }
                else
                {
                    chartObject1 = ChartObjects.Add(x1, 20 + (200 * globalCounter), 400, 350);
                }

                globalCounter++;

                Chart chart1 = chartObject1.Chart;
                chart1.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlXYScatterLines;
                chart1.HasTitle = true;
                chart1.ChartTitle.Text = "Cook's D Chart";
                chart1.HasLegend = false;
                seriesCollectionX = (SeriesCollection)chart1.SeriesCollection();
                seriesX = seriesCollectionX.NewSeries();
                seriesX.Values = model.CooksD.ToArray();
                seriesX.XValues = model.observationNumber.ToArray();
                //seriesX.Name = "Residuals";
                seriesX.MarkerStyle = XlMarkerStyle.xlMarkerStyleCircle;
                chart1.WallsAndGridlines2D = false;
                Axis axis1 = (Axis)chart1.Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary);
                axis1.HasTitle = true;
                axis1.AxisTitle.Text = "Cook's D";
                axis1.HasMajorGridlines = false;
                axis1.HasMinorGridlines = false;
                Axis axis21 = (Axis)chart1.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
                axis21.HasTitle = true;
                axis21.AxisTitle.Text = "Observation Number";
                axis21.HasMajorGridlines = false;
                axis21.HasMinorGridlines = false;
                axis21.MinimumScale = model.observationNumber.Column(0).Min();
                axis1.MinimumScale = model.CooksD.Column(0).Min();
                axis1.CrossesAt = model.CooksD.Column(0).Min();
                axis21.CrossesAt = model.observationNumber.Column(0).Min();
            }

            if (model.isDFFITSCheckedInPAndGSection)
            {
                i++;
                Microsoft.Office.Interop.Excel.ChartObject chartObject1;

                if ((i + 1) % 2 == 0)
                {
                    chartObject1 = ChartObjects.Add(x3, 20 + (200 * (globalCounter - ((globalCounter) % 2))), 400, 350);
                }
                else
                {
                    chartObject1 = ChartObjects.Add(x1, 20 + (200 * globalCounter), 400, 350);
                }

                globalCounter++;

                Chart chart1 = chartObject1.Chart;
                chart1.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlXYScatterLines;
                chart1.HasTitle = true;
                chart1.ChartTitle.Text = "DFFITS Chart";
                chart1.HasLegend = false;
                seriesCollectionX = (SeriesCollection)chart1.SeriesCollection();
                seriesX = seriesCollectionX.NewSeries();
                seriesX.Values = model.DFFITS.ToArray();
                seriesX.XValues = model.observationNumber.ToArray();
                //seriesX.Name = "Residuals";
                seriesX.MarkerStyle = XlMarkerStyle.xlMarkerStyleCircle;
                chart1.WallsAndGridlines2D = false;
                Axis axis1 = (Axis)chart1.Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary);
                axis1.HasTitle = true;
                axis1.AxisTitle.Text = "DFFITS";
                axis1.HasMajorGridlines = false;
                axis1.HasMinorGridlines = false;
                Axis axis21 = (Axis)chart1.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
                axis21.HasTitle = true;
                axis21.AxisTitle.Text = "Observation Number";
                axis21.HasMajorGridlines = false;
                axis21.HasMinorGridlines = false;

                axis21.MinimumScale = model.observationNumber.Column(0).Min();
                axis1.MinimumScale = model.DFFITS.Column(0).Min();
                axis1.CrossesAt = model.DFFITS.Column(0).Min();
                axis21.CrossesAt = model.observationNumber.Column(0).Min();
            }


        }

        public void createOutputOnASeparateSheet()
        {
            //ThisAddIn.form.updateStatus("Printing output on a new sheet ....");

            newWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets.Add();
            newWorksheet.Select();
            newWorksheet = ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A1", "A1").Value2 = "LINEAR REGRESSION OUTPUT";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A1", "A1").Cells.Font.Bold = true;

            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A2", "A2").Value2 = "Dependent Variable:";
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("B2", "B2").Value2 = model.yVariable;

            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A3", "B3").Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A1", "Z1").Cells.EntireColumn.ColumnWidth = 17;

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
                ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("I18", "I18").Value2 = "Standardized Coeff.";
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

            if (model.isConfidenceLimitsEnabledInAdvancedOptions || model.isOriginalEnabledInAdvancedOptions || model.isPredictedEnabledInAdvancedOptions || model.isResidualsEnabledInAdvancedOptions || model.isStandardizedResidualsEnabledInAdvancedOtions || model.isStudentizedResidualsEnabledInAdvancedOptions || model.isPRESSResidualsEnabledInAdvancedOptions || model.isRStudentEnabledInAdvancedOptions || model.isLeverageEnabledInAdvancedOptions || model.isCooksDEnabledInAdvancedOptions || model.isDFFITSEnabledInAdvancedOptions) {

                ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A" + advancedTableReference, "A" + advancedTableReference).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A" + (advancedTableReference + 1), "A" + (advancedTableReference + 1)).Value2 = "Obs #";
                ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A" + (advancedTableReference + 1), "A" + (advancedTableReference + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;

                int loopLimit = model.arrayYOriginalCopy.Length;

                if (model.isLabelsCheckedInBasic)
                {
                    loopLimit--;
                }

                for (int i = 0; i < loopLimit; i++)
                {
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A" + (advancedTableReference + 2 + i), "A" + (advancedTableReference + 2 + i)).Value2 = (i + 1);
                }

                ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range("A" + (advancedTableReference + loopLimit + 1), "A" + (advancedTableReference + loopLimit + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
            }

            if (model.isOriginalEnabledInAdvancedOptions)
            {
                advancedTableIndex++;
                ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + advancedTableReference, Util.IntToLetters(advancedTableIndex) + advancedTableReference).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1)).Value2 = model.yVariable;
                ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;

                int c = 0, r = 0, loopLimit = model.arrayYOriginalCopy.Length;

                if (model.isLabelsCheckedInBasic)
                {
                    loopLimit--;
                }

                for (int i = 0; i < loopLimit; i++)
                {
                    if (model.arrayXNoYRowNumbers.Contains(i+1) && model.isLabelsCheckedInBasic || model.arrayXNoYRowNumbers.Contains(i) && !model.isLabelsCheckedInBasic)
                    {
                        ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i)).Cells.Interior.Color = XlRgbColor.rgbLightPink;
                        ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i)).Value2 = "";
                        continue;
                    }                    
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i)).Value2 = model.arrayYConverted[r++, 0];
                }

                ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + loopLimit + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + loopLimit + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;

                int correction = 1;
                if (model.noIntercept)
                {
                    correction = 0;
                }

                c = 0; r = 0; loopLimit = model.arrayXOriginalCopy.GetLength(0);
                int startRow = 0;

                if (model.isLabelsCheckedInBasic)
                {
                    loopLimit--;
                    startRow = 1;
                }

                for (int i = 0; i < model.arrayXOriginalCopy.GetLength(1); i++)
                {
                    advancedTableIndex++;
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + advancedTableReference, Util.IntToLetters(advancedTableIndex) + advancedTableReference).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1)).Value2 = model.xVariables.ElementAt(i);
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;

                    for (int j = 0; j < loopLimit; j++)
                    {
                            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + j), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + j)).Value2 = model.arrayXOriginalCopy[j + startRow, i];
                    }

                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + loopLimit + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + loopLimit + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                }
            }

            if (model.isPredictedEnabledInAdvancedOptions)
            {
                int r = 0, c = 0, loopLimit = model.arrayYOriginalCopy.Length;

                if (model.isLabelsCheckedInBasic)
                {
                    loopLimit--;
                }

                advancedTableIndex++;
                for (int i = 0; i < loopLimit; i++)
                {
                    if (model.arrayXNoYRowNumbers.Contains(i + 1) && model.isLabelsCheckedInBasic || model.arrayXNoYRowNumbers.Contains(i) && !model.isLabelsCheckedInBasic)
                    {
                        ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i)).Cells.Interior.Color = XlRgbColor.rgbLightPink;
                        ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i)).Value2 = model.arrayXNoYComputedY[c++, 0];
                        continue;
                    }

                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + advancedTableReference, Util.IntToLetters(advancedTableIndex) + advancedTableReference).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1)).Value2 = "Predicted";
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i)).Value2 = model.yCap[r++, 0];
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + loopLimit + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + loopLimit + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                }

            }


            if (model.isConfidenceLimitsEnabledInAdvancedOptions)
            {
                int r = 0, c = 0, loopLimit = model.arrayYOriginalCopy.Length;

                if (model.isLabelsCheckedInBasic)
                {
                    loopLimit--;
                }

                advancedTableIndex++;
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + advancedTableReference, Util.IntToLetters(advancedTableIndex + 1) + advancedTableReference).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1)).Value2 = model.confidenceLevel + "% Mean LL";
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex + 1) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex + 1) + (advancedTableReference + 1)).Value2 = model.confidenceLevel + "% Mean UL";

                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex + 1) + (advancedTableReference + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                    for (int j = 0; j < loopLimit; j++)
                    {
                        if (model.arrayXNoYRowNumbers.Contains(j + 1) && model.isLabelsCheckedInBasic || model.arrayXNoYRowNumbers.Contains(j) && !model.isLabelsCheckedInBasic)
                        {
                        ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + j), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + j)).Value2 = model.LLMean[r++, 0]; r--;
                        ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex + 1) + (advancedTableReference + 2 + j), Util.IntToLetters(advancedTableIndex + 1) + (advancedTableReference + 2 + j)).Value2 = model.ULMean[r++, 0];
                        continue;
                        }

                        ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + j), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + j)).Value2 = model.LLMean[c++, 0]; c--;
                        ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex + 1) + (advancedTableReference + 2 + j), Util.IntToLetters(advancedTableIndex + 1) + (advancedTableReference + 2 + j)).Value2 = model.ULMean[c++, 0];
                    }
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + loopLimit + 1), Util.IntToLetters(advancedTableIndex + 1) + (advancedTableReference + loopLimit + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;

                advancedTableIndex++;
                advancedTableIndex++;
                r = 0; c = 0;

                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + advancedTableReference, Util.IntToLetters(advancedTableIndex + 1) + advancedTableReference).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1)).Value2 = model.confidenceLevel + "% Pred LL";
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex + 1) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex + 1) + (advancedTableReference + 1)).Value2 = model.confidenceLevel + "% Pred UL";

                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex + 1) + (advancedTableReference + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                    for (int j = 0; j < loopLimit; j++)
                    {
                        if (model.arrayXNoYRowNumbers.Contains(j + 1) && model.isLabelsCheckedInBasic || model.arrayXNoYRowNumbers.Contains(j) && !model.isLabelsCheckedInBasic)
                        {
                            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + j), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + j)).Value2 = model.LLPred[r++, 0]; r--;
                            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex + 1) + (advancedTableReference + 2 + j), Util.IntToLetters(advancedTableIndex + 1) + (advancedTableReference + 2 + j)).Value2 = model.ULPred[r++, 0];
                            continue;
                        }

                        ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + j), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + j)).Value2 = model.LLPred[c++, 0]; c--;
                        ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex + 1) + (advancedTableReference + 2 + j), Util.IntToLetters(advancedTableIndex + 1) + (advancedTableReference + 2 + j)).Value2 = model.ULPred[c++, 0];
                    }
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + loopLimit + 1), Util.IntToLetters(advancedTableIndex + 1) + (advancedTableReference + loopLimit + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;

                advancedTableIndex++;
            }

            if (model.isResidualsEnabledInAdvancedOptions)
            {
                int r=0, loopLimit = model.arrayYOriginalCopy.Length;

                if (model.isLabelsCheckedInBasic)
                {
                    loopLimit--;
                }

                advancedTableIndex++;
                for (int i = 0; i < loopLimit; i++)
                {
                    if (model.arrayXNoYRowNumbers.Contains(i + 1) && model.isLabelsCheckedInBasic || model.arrayXNoYRowNumbers.Contains(i) && !model.isLabelsCheckedInBasic)
                    {
                        ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i)).Value2 = "";
                        continue;
                    }

                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + advancedTableReference, Util.IntToLetters(advancedTableIndex) + advancedTableReference).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1)).Value2 = "Residual";
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i)).Value2 = model.residuals[r++, 0];
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + loopLimit + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + loopLimit + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                }
            }

            if (model.isStandardizedResidualsEnabledInAdvancedOtions)
            {
                int r=0, loopLimit = model.arrayYOriginalCopy.Length;

                if (model.isLabelsCheckedInBasic)
                {
                    loopLimit--;
                }

                advancedTableIndex++;
                for (int i = 0; i < loopLimit; i++)
                {
                    if (model.arrayXNoYRowNumbers.Contains(i + 1) && model.isLabelsCheckedInBasic || model.arrayXNoYRowNumbers.Contains(i) && !model.isLabelsCheckedInBasic)
                    {
                        ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i)).Value2 = "";
                        continue;
                    }

                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + advancedTableReference, Util.IntToLetters(advancedTableIndex) + advancedTableReference).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1)).Value2 = "Std. Residual";
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i)).Value2 = model.standardizedResiduals[r++, 0];
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + loopLimit + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + loopLimit + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                }

            }

            if (model.isStudentizedResidualsEnabledInAdvancedOptions)
            {
                int r = 0, loopLimit = model.arrayYOriginalCopy.Length;

                if (model.isLabelsCheckedInBasic)
                {
                    loopLimit--;
                }

                advancedTableIndex++;
                for (int i = 0; i < loopLimit; i++)
                {
                    if (model.arrayXNoYRowNumbers.Contains(i + 1) && model.isLabelsCheckedInBasic || model.arrayXNoYRowNumbers.Contains(i) && !model.isLabelsCheckedInBasic)
                    {
                        ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i)).Value2 = "";
                        continue;
                    }

                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + advancedTableReference, Util.IntToLetters(advancedTableIndex) + advancedTableReference).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1)).Value2 = "Studentized Residuals";
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i)).Value2 = model.studentizedResiduals[r++, 0];
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + loopLimit + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + loopLimit + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                }

            }

            if (model.isPRESSResidualsEnabledInAdvancedOptions)
            {
                int r = 0, loopLimit = model.arrayYOriginalCopy.Length;

                if (model.isLabelsCheckedInBasic)
                {
                    loopLimit--;
                }

                advancedTableIndex++;
                for (int i = 0; i < loopLimit; i++)
                {
                    if (model.arrayXNoYRowNumbers.Contains(i + 1) && model.isLabelsCheckedInBasic || model.arrayXNoYRowNumbers.Contains(i) && !model.isLabelsCheckedInBasic)
                    {
                        ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i)).Value2 = "";
                        continue;
                    }

                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + advancedTableReference, Util.IntToLetters(advancedTableIndex) + advancedTableReference).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1)).Value2 = "PRESS Residuals";
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i)).Value2 = model.PRESSResiduals[r++, 0];
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + loopLimit + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + loopLimit + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                }

            }

            if (model.isRStudentEnabledInAdvancedOptions)
            {
                int r = 0, loopLimit = model.arrayYOriginalCopy.Length;

                if (model.isLabelsCheckedInBasic)
                {
                    loopLimit--;
                }

                advancedTableIndex++;
                for (int i = 0; i < loopLimit; i++)
                {
                    if (model.arrayXNoYRowNumbers.Contains(i + 1) && model.isLabelsCheckedInBasic || model.arrayXNoYRowNumbers.Contains(i) && !model.isLabelsCheckedInBasic)
                    {
                        ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i)).Value2 = "";
                        continue;
                    }

                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + advancedTableReference, Util.IntToLetters(advancedTableIndex) + advancedTableReference).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1)).Value2 = "R-Student";
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i)).Value2 = model.RStudentResiduals[r++, 0];
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + loopLimit + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + loopLimit + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                }

            }

            if (model.isLeverageEnabledInAdvancedOptions)
            {
                int r = 0, loopLimit = model.arrayYOriginalCopy.Length;

                if (model.isLabelsCheckedInBasic)
                {
                    loopLimit--;
                }

                advancedTableIndex++;
                for (int i = 0; i < loopLimit; i++)
                {
                    if (model.arrayXNoYRowNumbers.Contains(i + 1) && model.isLabelsCheckedInBasic || model.arrayXNoYRowNumbers.Contains(i) && !model.isLabelsCheckedInBasic)
                    {
                        ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i)).Value2 = "";
                        continue;
                    }

                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + advancedTableReference, Util.IntToLetters(advancedTableIndex) + advancedTableReference).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1)).Value2 = "Leverage";
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i)).Value2 = model.Leverage[r++, 0];
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + loopLimit + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + loopLimit + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                }

            }

            if (model.isCooksDEnabledInAdvancedOptions)
            {
                int r = 0, loopLimit = model.arrayYOriginalCopy.Length;

                if (model.isLabelsCheckedInBasic)
                {
                    loopLimit--;
                }

                advancedTableIndex++;
                for (int i = 0; i < loopLimit; i++)
                {
                    if (model.arrayXNoYRowNumbers.Contains(i + 1) && model.isLabelsCheckedInBasic || model.arrayXNoYRowNumbers.Contains(i) && !model.isLabelsCheckedInBasic)
                    {
                        ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i)).Value2 = "";
                        continue;
                    }

                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + advancedTableReference, Util.IntToLetters(advancedTableIndex) + advancedTableReference).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1)).Value2 = "Cook's D";
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i)).Value2 = model.CooksD[r++, 0];
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + loopLimit + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + loopLimit + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                }

            }

            if (model.isDFFITSEnabledInAdvancedOptions)
            {
                int r = 0, loopLimit = model.arrayYOriginalCopy.Length;

                if (model.isLabelsCheckedInBasic)
                {
                    loopLimit--;
                }

                advancedTableIndex++;
                for (int i = 0; i < loopLimit; i++)
                {
                    if (model.arrayXNoYRowNumbers.Contains(i + 1) && model.isLabelsCheckedInBasic || model.arrayXNoYRowNumbers.Contains(i) && !model.isLabelsCheckedInBasic)
                    {
                        ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i)).Value2 = "";
                        continue;
                    }

                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + advancedTableReference, Util.IntToLetters(advancedTableIndex) + advancedTableReference).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1)).Value2 = "DFFITS";
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + 2 + i)).Value2 = model.DFFITS[r++, 0];
                    ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex) + (advancedTableReference + loopLimit + 1), Util.IntToLetters(advancedTableIndex) + (advancedTableReference + loopLimit + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                }

            }


            if (model.isDFBETASEnabledInAdvancedOptions)
            {
                int r = 0, loopLimit = model.arrayYOriginalCopy.Length;

                if (model.isLabelsCheckedInBasic)
                {
                    loopLimit--;
                }

                advancedTableIndex++;
                for (int i = 0; i < loopLimit; i++)
                {
                    for (int j = 0; j < model.DFBETAS.ColumnCount; j++)
                    {
                        if (model.arrayXNoYRowNumbers.Contains(i + 1) && model.isLabelsCheckedInBasic || model.arrayXNoYRowNumbers.Contains(i) && !model.isLabelsCheckedInBasic)
                        {
                            ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex + j) + (advancedTableReference + 2 + i), Util.IntToLetters(advancedTableIndex + j) + (advancedTableReference + 2 + i)).Value2 = "";
                            continue;
                        }

                        ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex + j) + advancedTableReference, Util.IntToLetters(advancedTableIndex + j) + advancedTableReference).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                        ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex + j) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex + j) + (advancedTableReference + 1)).Value2 = "DFBETA " + j;
                        ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex + j) + (advancedTableReference + 1), Util.IntToLetters(advancedTableIndex + j) + (advancedTableReference + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                        ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex + j) + (advancedTableReference + 2 + i), Util.IntToLetters(advancedTableIndex + j) + (advancedTableReference + 2 + i)).Value2 = model.DFBETAS[r, j];
                        ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(Util.IntToLetters(advancedTableIndex + j) + (advancedTableReference + loopLimit + 1), Util.IntToLetters(advancedTableIndex + j) + (advancedTableReference + loopLimit + 1)).Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;

                    }
                    if (model.arrayXNoYRowNumbers.Contains(i + 1) && model.isLabelsCheckedInBasic || model.arrayXNoYRowNumbers.Contains(i) && !model.isLabelsCheckedInBasic)
                    {
                        r--;
                    }

                    r++;
                }
            }

        }
    }
}
