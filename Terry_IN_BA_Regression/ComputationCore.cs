using MathNet.Numerics.LinearAlgebra;
using MathNet.Numerics.LinearAlgebra.Double;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Terry_IN_BA_Regression
{
    class ComputationCore
    {
        InputModel input;
        OutputModel output;

        public ComputationCore(InputModel input, OutputModel output)
        {
            this.input = input;
            this.output = output;
            adjustRowsForNullEntries();
            init();
            doCompute();
        }

        public InputModel getInputModel()
        {
            return this.input;
        } 

        public OutputModel getOutputModel()
        {
            return this.output;
        }

        private void doCompute()
        {
            computeBetaCoefficients();
            computeK();
            computeDfe();
            computeYCap();
            computeYBar();
            computeSSR();
            computeSSE();
            computeSST();
            computeMSR();
            computeMSE();
            computeFStatistic();
            computePValue();
            computeRSquare();
            computeR();
            computeRSQAdjusted();
            computeSE();
            initializeMatrices();
            computeStandardizedCoefficients();
            computeTStatistic();
            computePValueForBetaCoefficients();
            computeTStar();
            computeLLAndUL();
            computeVIF();
            bundleCompute();
            computeYCapforOddTuples();

            if (output.arrayXNoYRowNumbers.Count != 0)
            {
                computeYforAllXNoY();
            }
        }

        private void init()
        {
            ThisAddIn.form.updateStatus("Parsing raw data  ....");
           

            double[,] doubleY;
            if (output.isLabelsCheckedInBasic)
            {
                doubleY = new double[output.yR - 1, output.yC];
                for (int x = 1; x < output.arrayY.GetLength(0); x++)
                {
                    for (int y = 0; y < output.arrayY.GetLength(1); y++)
                    {
                        double.TryParse(output.arrayY[x, y], out doubleY[x - 1, y]);
                    }
                }
            }
            else
            {
                doubleY = new double[output.yR, output.yC];
                for (int x = 0; x < output.arrayY.GetLength(0); x++)
                {
                    for (int y = 0; y < output.arrayY.GetLength(1); y++)
                    {
                        double.TryParse(output.arrayY[x, y], out doubleY[x, y]);
                    }
                }
            }
            output.arrayYConverted = DenseMatrix.OfArray(doubleY);
            output.n = doubleY.Length;

            double[,] doubleX = { };

            if (output.noIntercept)
            {
                if (output.isLabelsCheckedInBasic)
                {
                    output.yVariable = output.arrayY[0, 0];
                    for (int y = 0; y < output.arrayX.GetLength(1); y++)
                    {
                        output.xVariables.AddLast(output.arrayX[0, y]);
                        output.xVariableStates.AddLast(true);
                    }
                    doubleX = new double[output.xR - 1, output.xC];
                    for (int x = 1; x < output.arrayX.GetLength(0); x++)
                    {
                        for (int y = 0; y < output.arrayX.GetLength(1); y++)
                        {
                            double.TryParse(output.arrayX[x, y], out doubleX[x - 1, y]);
                        }
                    }
                }
                else
                {
                    output.yVariable = "Y";
                    doubleX = new double[output.xR, output.xC];
                    for (int x = 0; x < output.arrayX.GetLength(0); x++)
                    {
                        for (int y = 0; y < output.arrayX.GetLength(1); y++)
                        {
                            double.TryParse(output.arrayX[x, y], out doubleX[x, y]);
                        }
                    }
                }
            }
            else
            {
                if (output.isLabelsCheckedInBasic)
                {
                    output.yVariable = output.arrayY[0, 0];
                    for (int y = 0; y < output.arrayX.GetLength(1); y++)
                    {
                        output.xVariables.AddLast(output.arrayX[0, y]);
                        output.xVariableStates.AddLast(true);
                    }

                    doubleX = new double[output.xR - 1, output.xC + 1];
                    for (int x = 0; x < (output.arrayX.GetLength(0) - 1); x++)
                    {
                        doubleX[x, 0] = 1;
                    }
                    for (int x = 1; x < output.arrayX.GetLength(0); x++)
                    {
                        for (int y = 0; y < output.arrayX.GetLength(1); y++)
                        {
                            double.TryParse(output.arrayX[x, y], out doubleX[x - 1, y + 1]);
                        }
                    }
                }
                else
                {
                    output.yVariable = "Y";
                    for (int y = 0; y < output.arrayX.GetLength(1); y++)
                    {
                        output.xVariables.AddLast("Variable X" + (y + 1));
                        output.xVariableStates.AddLast(true);
                    }

                    doubleX = new double[output.xR, output.xC + 1];
                    for (int x = 0; x < output.arrayX.GetLength(0); x++)
                    {
                        doubleX[x, 0] = 1;
                    }
                    for (int x = 0; x < output.arrayX.GetLength(0); x++)
                    {
                        for (int y = 0; y < output.arrayX.GetLength(1); y++)
                        {
                            double.TryParse(output.arrayX[x, y], out doubleX[x, y + 1]);
                        }
                    }
                }
            }
            output.arrayXConverted = DenseMatrix.OfArray(doubleX);
     
            
            if (output.noIntercept)
            {
                for (int i = output.xVariableStates.Count - 1; i >= 0; i--)
                {
                    if (!output.xVariableStates.ElementAt(i))
                    {
                        output.arrayXConverted = output.arrayXConverted.RemoveColumn(i);
                        output.xC--;
                        output.xVariables.Remove(output.xVariables.ElementAt(i));
                    }
                }
            }
            else
            {
                for (int i = output.xVariableStates.Count - 1; i >= 0; i--)
                {
                    if (!output.xVariableStates.ElementAt(i))
                    {
                        output.arrayXConverted = output.arrayXConverted.RemoveColumn(i + 1);
                        output.xC--;
                        output.xVariables.Remove(output.xVariables.ElementAt(i));
                    }
                }
            }

            ThisAddIn.form.updateStatus("Parsing raw data  .... [OK]");

        }

        private void adjustRowsForNullEntries()
        {  
            output.arrayXOriginalCopy = output.arrayX;
            output.arrayYOriginalCopy = output.arrayY;

            ThisAddIn.form.updateStatus("Filtering odd tuples  ....");

            if (output.isLabelsCheckedInBasic)
            {
                output.originalRows = output.yR - 1;
            }
            else
            {
                output.originalRows = output.yR;
            }

            int i, j;
            HashSet<int> removeRowSet = new HashSet<int>();

            for (i = 0; i < output.xR; i++)
            {
                int count = 0;
                for (j = 0; j < output.xC; j++)
                {
                    if (output.arrayX[i, j] == "")
                    {
                        removeRowSet.Add(i);
                    } else
                    {
                        ++count;
                    }

                    if (output.arrayY[i, 0] == "")
                    {
                        removeRowSet.Add(i);
                    }
                }

                if (count == output.xC && output.arrayY[i, 0] == "")
                {
                        output.arrayXNoYRowNumbers.Add(i);
                }

            }

            if (output.arrayXNoYRowNumbers.Count != 0) {

                double[,] p;

                if (output.noIntercept)
                {
                    p = new double[output.arrayXNoYRowNumbers.Count, output.arrayXOriginalCopy.GetLength(1)];
                }
                else
                {
                    p = new double[output.arrayXNoYRowNumbers.Count, output.arrayXOriginalCopy.GetLength(1) + 1];
                }

                for (int a = 0; a < output.arrayXNoYRowNumbers.Count; a++)
                {
                    if (output.noIntercept)
                    {
                        for (int b = 0; b < output.arrayXOriginalCopy.GetLength(1); b++)
                        {
                            p[a, b] = 1;
                        }
                    } else
                    {
                        for (int b = 0; b < (output.arrayXOriginalCopy.GetLength(1) + 1); b++)
                        {
                            p[a, b] = 1;
                        }
                    }
                }

                output.arrayXNoY = DenseMatrix.OfArray(p);
                int r = 0;
                foreach (int value in output.arrayXNoYRowNumbers)
                {
                    for (int l = 0; l < output.arrayXOriginalCopy.GetLength(1); l++)
                    {
                        double number;
                        double.TryParse(output.arrayXOriginalCopy[value, l], out number);
                        if (output.noIntercept)
                        {
                            output.arrayXNoY[r, l] = number;
                        } else
                        {
                            output.arrayXNoY[r, l+1] = number;
                        }
                    }
                    r++;
                }

            }


            output.arrayXFilter = new string[output.xR - removeRowSet.Count, output.xC];

            int x = 0, y = 0;
            for (i = 0; i < output.xR; i++)
            {
                if (!removeRowSet.Contains(i))
                {
                    y = 0;
                    for (j = 0; j < output.xC; j++)
                    {
                        output.arrayXFilter[x, y] = output.arrayX[i, j];
                        y++;
                    }
                    x++;
                }
            }
            output.arrayX = output.arrayXFilter;

            output.arrayYFilter = new string[output.yR - removeRowSet.Count, output.yC];

            x = 0; y = 0;
            for (i = 0; i < output.yR; i++)
            {
                if (!removeRowSet.Contains(i))
                {
                    if (output.arrayY[i, 0] == "")
                    {

                    }
                    else
                    {
                        y = 0;
                        for (j = 0; j < output.yC; j++)
                        {
                            output.arrayYFilter[x, y] = output.arrayY[i, j];
                            y++;
                        }
                        x++;
                    }
                }
            }
            output.arrayY = output.arrayYFilter;
            output.xR = output.yR = output.arrayYFilter.Length;
            output.n = output.arrayYFilter.Length;

            ThisAddIn.form.updateStatus("Filtering odd tuples  .... [OK]");

        }

        private void computeBetaCoefficients()
        {
            ThisAddIn.form.updateStatus("Computing coefficients  ....");

            output.betaCoefficients = ((output.arrayXConverted.Transpose().Multiply(output.arrayXConverted)).Inverse()).Multiply(output.arrayXConverted.Transpose()).Multiply(output.arrayYConverted);
        }

        private void computeK()
        {
            ThisAddIn.form.updateStatus("Computing K  ....");

            output.k = output.xVariables.Count;
        }

        private void computeDfe()
        {
            ThisAddIn.form.updateStatus("Computing DFE  ....");

            if (output.noIntercept)
            {
                output.dfe = output.n - output.xVariables.Count;
            }
            else
            {
                output.dfe = output.n - output.xVariables.Count - 1;
            }
        }

        private void computeYCap()
        {
            ThisAddIn.form.updateStatus("Computing Y hat  ....");

            output.yCap = output.arrayXConverted.Multiply(output.betaCoefficients);
        }

        private void computeYBar()
        {
            ThisAddIn.form.updateStatus("Computing Y bar  ....");

            double ySum = 0;
            for (int i = 0; i < output.arrayYConverted.RowCount; i++)
            {
                ySum += output.arrayYConverted.At(i, 0);
            }
            output.yBar = ySum / output.arrayYConverted.RowCount;
        }

        private void computeSSR()
        {
            ThisAddIn.form.updateStatus("Computing SSR  ....");

            for (int i = 0; i < output.yCap.RowCount; i++)
            {
                output.SSR += (output.yCap.At(i, 0) - output.yBar) * (output.yCap.At(i, 0) - output.yBar);
            }
        }

        private void computeSSE()
        {
            ThisAddIn.form.updateStatus("Computing SSE  ....");

            for (int i = 0; i < output.yCap.RowCount; i++)
            {
                output.SSE += (output.arrayYConverted.At(i, 0) - output.yCap.At(i, 0)) * (output.arrayYConverted.At(i, 0) - output.yCap.At(i, 0));
            }
        }

        private void computeSST()
        {
            ThisAddIn.form.updateStatus("Computing SST  ....");

            for (int i = 0; i < output.yCap.RowCount; i++)
            {
                output.SST += (output.arrayYConverted.At(i, 0) - output.yBar) * (output.arrayYConverted.At(i, 0) - output.yBar);
            }
        }

        private void computeMSR()
        {
            ThisAddIn.form.updateStatus("Computing MSR  ....");

            output.MSR = output.SSR / output.k;
        }

        private void computeMSE()
        {
            ThisAddIn.form.updateStatus("Computing MSE  ....");

            output.MSE = output.SSE / output.dfe;
        }

        private void computeFStatistic()
        {
            ThisAddIn.form.updateStatus("Computing F Statistic  ....");

            output.F = output.MSR / output.MSE;
        }

        private void computePValue()
        {
            ThisAddIn.form.updateStatus("Computing P value  ....");

            output.P = MathNet.Numerics.ExcelFunctions.FDist(output.F, (int)output.k, (int)(output.dfe));
        }

        private void computeRSquare()
        {
            ThisAddIn.form.updateStatus("Computing R Square  ....");

            output.RSquare = output.SSR / output.SST;
        }

        private void computeR()
        {
            ThisAddIn.form.updateStatus("Computing R  ....");

            output.R = Math.Sqrt(output.RSquare);
        }

        private void computeRSQAdjusted()
        {
            ThisAddIn.form.updateStatus("Computing R square adjusted ....");

            output.RSqAdj = (((output.n - 1) * output.RSquare) - output.k) / (output.dfe);
        }

        private void computeSE()
        {
            ThisAddIn.form.updateStatus("Computing standard error ....");

            output.SE = output.MSE * output.arrayXConverted.Transpose().Multiply(output.arrayXConverted).Inverse();

            for (int i = 0; i < output.SE.RowCount; i++)
            {
                output.SE[i, 0] = Math.Sqrt(output.SE.At(i, i));
            }
        }

        private void initializeMatrices()
        {
            ThisAddIn.form.updateStatus("setting matrices ....");

            double[,] p = new double[output.betaCoefficients.RowCount, 1];

            for (int i = 0; i < output.xVariables.Count; i++)
            {
                p[i, 0] = 1;
            }

            output.tStat = DenseMatrix.OfArray(p);
            output.pValue = DenseMatrix.OfArray(p);
            output.LL = DenseMatrix.OfArray(p);
            output.UL = DenseMatrix.OfArray(p);
            output.standardizedCoefficients = DenseMatrix.OfArray(p);


            p = new double[output.arrayXConverted.RowCount, 1];

            for (int i = 0; i < output.arrayXConverted.RowCount; i++)
            {
                p[i, 0] = 1;
            }

            output.LLMean = DenseMatrix.OfArray(p);
            output.ULMean = DenseMatrix.OfArray(p);
            output.LLPred = DenseMatrix.OfArray(p);
            output.ULPred = DenseMatrix.OfArray(p);
            output.residuals = DenseMatrix.OfArray(p);
            output.standardizedResiduals = DenseMatrix.OfArray(p);
            output.studentizedResiduals = DenseMatrix.OfArray(p);
            output.PRESSResiduals = DenseMatrix.OfArray(p);
            output.RStudentResiduals = DenseMatrix.OfArray(p);
            output.Leverage = DenseMatrix.OfArray(p);
            output.CooksD = DenseMatrix.OfArray(p);
            output.DFFITS = DenseMatrix.OfArray(p);

            p = new double[1, output.xVariables.Count];

            for (int i = 0; i < output.xVariables.Count; i++)
            {
                p[0, i] = 1;
            }

            output.SDx = DenseMatrix.OfArray(p);


            p = new double[output.arrayXConverted.RowCount, output.arrayXConverted.ColumnCount];

            for (int i = 0; i < output.arrayXConverted.RowCount; i++)
            {
                for (int j = 0; j < output.arrayXConverted.ColumnCount; j++)
                {
                    p[i, j] = 1;
                }
            }

            output.DFBETAS = DenseMatrix.OfArray(p);

            if (output.arrayXNoYRowNumbers.Count != 0) {
                p = new double[output.arrayXNoY.RowCount, 1];

                for (int i = 0; i < output.arrayXNoY.RowCount; i++)
                {
                    p[i, 0] = 1;
                }

                output.arrayXNoYLLMean = DenseMatrix.OfArray(p);
                output.arrayXNoYULMean = DenseMatrix.OfArray(p);
                output.arrayXNoYLLPred = DenseMatrix.OfArray(p);
                output.arrayXNoYULPred = DenseMatrix.OfArray(p);
            }
        }

        private void computeStandardizedCoefficients()
        {
            ThisAddIn.form.updateStatus("Computing standardized coefficients ....");

            if (output.noIntercept)
            {
                for (int i = 0; i < output.xVariables.Count; i++)
                {
                    output.SDx[0, i] = MathNet.Numerics.Statistics.Statistics.StandardDeviation(output.arrayXConverted.Column(i));
                }
            }
            else
            {
                for (int i = 0; i < output.xVariables.Count; i++)
                {
                    output.SDx[0, i] = MathNet.Numerics.Statistics.Statistics.StandardDeviation(output.arrayXConverted.Column(i + 1));
                }
            }

            output.SDy = MathNet.Numerics.Statistics.Statistics.StandardDeviation(output.arrayYConverted.Column(0));

            if (output.noIntercept)
            {
                for (int i = 0; i < output.xVariables.Count; i++)
                {
                    output.standardizedCoefficients[i, 0] = output.betaCoefficients.At(i, 0) * (output.SDx.At(0, i) / output.SDy);
                }
            }
            else
            {
                for (int i = 0; i < output.xVariables.Count; i++)
                {
                    output.standardizedCoefficients[i, 0] = output.betaCoefficients.At(i + 1, 0) * (output.SDx.At(0, i) / output.SDy);
                }
            }
        }

        private void computeTStatistic()
        {
            ThisAddIn.form.updateStatus("Computing T Statistic ....");

            for (int i = 0; i < output.SE.RowCount; i++)
            {
                output.tStat[i, 0] = output.betaCoefficients.At(i, 0) / output.SE.At(i, 0);
            }
        }

        private void computePValueForBetaCoefficients()
        {
            ThisAddIn.form.updateStatus("Computing P values for beta coefficients ....");

            for (int i = 0; i < output.tStat.RowCount; i++)
            {
                if (output.tStat[i, 0] <= 0)
                {
                    output.pValue[i, 0] = 2 - 2 * MathNet.Numerics.ExcelFunctions.TDist(output.tStat[i, 0], (int)output.dfe, 1);
                }
                else if (output.tStat[i, 0] > 0)
                {
                    output.pValue[i, 0] = 2 - 2 * (1 - MathNet.Numerics.ExcelFunctions.TDist(output.tStat[i, 0], (int)output.dfe, 1));
                }
            }
        }

        private void computeTStar()
        {
            ThisAddIn.form.updateStatus("Computing T* ....");

            output.tStar = Math.Abs(MathNet.Numerics.Distributions.StudentT.InvCDF(0d, 1d, (int)output.dfe, (1 - (double.Parse(output.confidenceLevel) / 100)) / 2));
        }

        private void computeLLAndUL()
        {
            ThisAddIn.form.updateStatus("Computing LL & UL ....");

            for (int i = 0; i < output.LL.RowCount; i++)
            {
                output.LL[i, 0] = output.betaCoefficients[i, 0] - output.tStar * output.SE[i, 0];
                output.UL[i, 0] = output.betaCoefficients[i, 0] + output.tStar * output.SE[i, 0];
            }
        }

        private void computeVIF()
        {
            ThisAddIn.form.updateStatus("Computing VIF ....");

            if (output.noIntercept)
            {
                double[,] w = new double[output.arrayXConverted.RowCount, output.arrayXConverted.ColumnCount];
                for (int i = 0; i < output.arrayXConverted.RowCount; i++)
                {
                    for (int j = 0; j < output.arrayXConverted.ColumnCount; j++)
                    {
                        w[i, j] = 1;
                    }
                }
                Matrix<double> W = DenseMatrix.OfArray(w);

                double[,] s = new double[1, output.arrayXConverted.ColumnCount];
                for (int i = 0; i < output.arrayXConverted.ColumnCount; i++)
                {
                    s[0, i] = 1;
                }
                Matrix<double> S = DenseMatrix.OfArray(s);

                double[,] xbar = new double[1, output.arrayXConverted.ColumnCount];
                for (int i = 0; i < output.arrayXConverted.ColumnCount; i++)
                {
                    xbar[0, i] = 1;
                }
                Matrix<double> xBar = DenseMatrix.OfArray(xbar);

                for (int i = 0; i < output.arrayXConverted.ColumnCount; i++)
                {
                    double xB = 0;
                    double sum = 0;
                    for (int j = 0; j < output.arrayXConverted.RowCount; j++)
                    {
                        sum += output.arrayXConverted.At(j, i);
                    }
                    xB = sum / output.arrayXConverted.RowCount;
                    xBar[0, i] = xB;

                    double sjj = 0;
                    for (int j = 0; j < output.arrayXConverted.RowCount; j++)
                    {
                        sjj += (output.arrayXConverted.At(j, i) - xB) * (output.arrayXConverted.At(j, i) - xB);
                    }
                    S[0, i] = sjj;
                }

                for (int i = 0; i < output.arrayXConverted.RowCount; i++)
                {
                    for (int j = 0; j < output.arrayXConverted.ColumnCount; j++)
                    {
                        W[i, j] = (output.arrayXConverted[i, j] - xBar[0, j]) / (Math.Sqrt(S[0, j]));
                    }
                }

                output.VIFMatrix = (W.Transpose().Multiply(W)).Inverse();
            }
            else
            {
                double[,] w = new double[output.arrayXConverted.RowCount, output.arrayXConverted.ColumnCount - 1];
                for (int i = 0; i < output.arrayXConverted.RowCount; i++)
                {
                    for (int j = 0; j < output.arrayXConverted.ColumnCount - 1; j++)
                    {
                        w[i, j] = 1;
                    }
                }
                Matrix<double> W = DenseMatrix.OfArray(w);

                double[,] s = new double[1, output.arrayXConverted.ColumnCount - 1];
                for (int i = 0; i < output.arrayXConverted.ColumnCount - 1; i++)
                {
                    s[0, i] = 1;
                }
                Matrix<double> S = DenseMatrix.OfArray(s);

                double[,] xbar = new double[1, output.arrayXConverted.ColumnCount - 1];
                for (int i = 0; i < output.arrayXConverted.ColumnCount - 1; i++)
                {
                    xbar[0, i] = 1;
                }
                Matrix<double> xBar = DenseMatrix.OfArray(xbar);

                for (int i = 1; i < output.arrayXConverted.ColumnCount; i++)
                {
                    double xB = 0;
                    double sum = 0;
                    for (int j = 0; j < output.arrayXConverted.RowCount; j++)
                    {
                        sum += output.arrayXConverted.At(j, i);
                    }
                    xB = sum / output.arrayXConverted.RowCount;
                    xBar[0, i - 1] = xB;

                    double sjj = 0;
                    for (int j = 0; j < output.arrayXConverted.RowCount; j++)
                    {
                        sjj += (output.arrayXConverted.At(j, i) - xB) * (output.arrayXConverted.At(j, i) - xB);
                    }
                    S[0, i - 1] = sjj;
                }

                for (int i = 0; i < output.arrayXConverted.RowCount; i++)
                {
                    for (int j = 1; j < output.arrayXConverted.ColumnCount; j++)
                    {
                        W[i, j - 1] = (output.arrayXConverted[i, j] - xBar[0, j - 1]) / (Math.Sqrt(S[0, j - 1]));
                    }
                }

                output.VIFMatrix = (W.Transpose().Multiply(W)).Inverse();
            }


        }

        private void bundleCompute()
        {
            Matrix<double> R = null;

            if (output.isDFBETASEnabledInAdvancedOptions)
            {
                R = (output.arrayXConverted.Transpose().Multiply(output.arrayXConverted)).Inverse().Multiply(output.arrayXConverted.Transpose());
            }

            for (int i = 0; i < output.arrayXConverted.RowCount; i++)
            {
                if (output.isConfidenceLimitsEnabledInAdvancedOptions)
                {
                    ThisAddIn.form.updateStatus("Computing LL & UL Mean and Predicted Vectors ....");

                    Matrix<double> m1 = output.arrayXConverted.Row(i).ToRowMatrix();
                    Matrix<double> m2 = output.arrayXConverted.Transpose();
                    Matrix<double> m3 = output.arrayXConverted;
                    Matrix<double> m4 = output.arrayXConverted.Row(i).ToRowMatrix().Transpose();

                    output.LLMean[i, 0] = output.yCap[i, 0] - output.tStar * Math.Sqrt(output.MSE * (m1.Multiply(((m2.Multiply(m3)).Inverse())).Multiply(m4)[0, 0]));
                    output.ULMean[i, 0] = output.yCap[i, 0] + output.tStar * Math.Sqrt(output.MSE * (m1.Multiply(((m2.Multiply(m3)).Inverse())).Multiply(m4)[0, 0]));

                    output.LLPred[i, 0] = output.yCap[i, 0] - output.tStar * Math.Sqrt(output.MSE * (1 + m1.Multiply(((m2.Multiply(m3)).Inverse())).Multiply(m4)[0, 0]));
                    output.ULPred[i, 0] = output.yCap[i, 0] + output.tStar * Math.Sqrt(output.MSE * (1 + m1.Multiply(((m2.Multiply(m3)).Inverse())).Multiply(m4)[0, 0]));
                }

                if (output.isResidualsEnabledInAdvancedOptions)
                {
                    ThisAddIn.form.updateStatus("Computing Residual Vectors ....");

                    output.residuals[i, 0] = output.arrayYConverted[i, 0] - output.yCap[i, 0];
                }

                output.residuals[i, 0] = output.arrayYConverted[i, 0] - output.yCap[i, 0];


                if (output.isStandardizedResidualsEnabledInAdvancedOtions)
                {
                    ThisAddIn.form.updateStatus("Computing Standardized Residual Vectors ....");

                    output.standardizedResiduals[i, 0] = (output.arrayYConverted[i, 0] - output.yCap[i, 0]) / (Math.Sqrt(output.MSE));
                }

                Matrix<double> H = null;

                if (output.isStudentizedResidualsEnabledInAdvancedOptions || output.isPRESSResidualsEnabledInAdvancedOptions || output.isRStudentEnabledInAdvancedOptions || output.isLeverageEnabledInAdvancedOptions || output.isCooksDEnabledInAdvancedOptions || output.isDFFITSEnabledInAdvancedOptions || output.isDFBETASEnabledInAdvancedOptions)
                {
                    ThisAddIn.form.updateStatus("Computing H ....");

                    H = output.arrayXConverted.Multiply((output.arrayXConverted.Transpose().Multiply(output.arrayXConverted)).Inverse()).Multiply(output.arrayXConverted.Transpose());
                }

                if (output.isRStudentEnabledInAdvancedOptions)
                {
                    output.studentizedResiduals[i, 0] = output.residuals[i, 0] / (Math.Sqrt(output.MSE * (1 - H[i, i])));
                }

                if (output.isPRESSResidualsEnabledInAdvancedOptions)
                {
                    ThisAddIn.form.updateStatus("Computing PRESS Residual Vectors ....");

                    output.PRESSResiduals[i, 0] = output.residuals[i, 0] / (1 - H[i, i]);
                }
              

                if (output.isRStudentEnabledInAdvancedOptions || output.isDFBETASEnabledInAdvancedOptions)
                {
                    ThisAddIn.form.updateStatus("Computing R-Student Vectors ....");

                    double sSquare = ((((output.n - output.k - 1) * output.MSE) - ((output.residuals[i, 0] * output.residuals[i, 0]) / (1 - H[i, i]))) / (output.n - output.k - 2));
                    output.RStudentResiduals[i, 0] = (output.residuals[i, 0] / (Math.Sqrt(sSquare * (1 - H[i, i]))));
                }


                if (output.isLeverageEnabledInAdvancedOptions)
                {
                    ThisAddIn.form.updateStatus("Computing Leverage Vectors ....");

                    output.Leverage[i, 0] = H[i, i];
                }

                if (output.isCooksDEnabledInAdvancedOptions)
                {
                    ThisAddIn.form.updateStatus("Computing Cooks D Vectors ....");

                    output.CooksD[i, 0] = (output.studentizedResiduals[i, 0] * output.studentizedResiduals[i, 0] * H[i, i]) / ((output.k + 1) * (1 - H[i, i]));
                }

                if (output.isDFFITSEnabledInAdvancedOptions)
                {
                    ThisAddIn.form.updateStatus("Computing DFFITS Vectors ....");

                    output.DFFITS[i, 0] = Math.Sqrt(H[i, i] / (1 - H[i, i])) * output.RStudentResiduals[i, 0];
                }

                if (output.isDFBETASEnabledInAdvancedOptions)
                {
                    ThisAddIn.form.updateStatus("Computing DFBETA Vectors ....");

                    for (int j = 0; j < output.arrayXConverted.ColumnCount; j++)
                        {
                            output.DFBETAS[i, j] = (R[j, i] / Math.Sqrt(R.Row(j).ToRowMatrix().Multiply(R.Row(j).ToColumnMatrix())[0, 0])) * (output.RStudentResiduals[i, 0] / (Math.Sqrt(1- H[i, i])));
                        }
                }
            }
        }

        public void computeYforAllXNoY()
        {
            output.arrayXNoYComputedY = output.arrayXNoY.Multiply(output.betaCoefficients);

            for (int i = 0; i < output.arrayXNoY.RowCount; i++)
            {
                Matrix<double> m1 = output.arrayXNoY.Row(i).ToRowMatrix();
                Matrix<double> m2 = output.arrayXConverted.Transpose();
                Matrix<double> m3 = output.arrayXConverted;
                Matrix<double> m4 = output.arrayXNoY.Row(i).ToRowMatrix().Transpose();

                output.arrayXNoYLLMean[i, 0] = output.arrayXNoYComputedY[i, 0] - output.tStar * Math.Sqrt(output.MSE * (m1.Multiply(((m2.Multiply(m3)).Inverse())).Multiply(m4)[0, 0]));
                output.arrayXNoYULMean[i, 0] = output.arrayXNoYComputedY[i, 0] + output.tStar * Math.Sqrt(output.MSE * (m1.Multiply(((m2.Multiply(m3)).Inverse())).Multiply(m4)[0, 0]));
                output.arrayXNoYLLPred[i, 0] = output.arrayXNoYComputedY[i, 0] - output.tStar * Math.Sqrt(output.MSE * (1 + m1.Multiply(((m2.Multiply(m3)).Inverse())).Multiply(m4)[0, 0]));
                output.arrayXNoYULPred[i, 0] = output.arrayXNoYComputedY[i, 0] + output.tStar * Math.Sqrt(output.MSE * (1 + m1.Multiply(((m2.Multiply(m3)).Inverse())).Multiply(m4)[0, 0]));
            }
        }

        public void clearCache()
        {
            output.arrayX = new string[0, 0];
            output.arrayY = new string[0, 0];
            output.arrayXFilter = new string[0, 0];
            output.arrayYFilter = new string[0, 0];
            output.xVariables = new LinkedList<string>();
            output.xVariableStates = new LinkedList<bool>();
            output.n = 0.0;
            output.k = 0.0;
            output.yR = 0;
            output.yC = 0;
            output.xR = 0;
            output.xC = 0;
            output.SSR = 0.0;
            output.yBar = 0.0;
            output.SSE = 0.0;
            output.SST = 0.0;
            output.MSR = 0.0;
            output.MSE = 0.0;
            output.F = 0.0;
            output.P = 0.0;
            output.RSquare = 0.0;
            output.R = 0.0;
            output.RSqAdj = 0.0;
            output.SDy = 0.0;
            output.dfe = 0.0;
            output.tStar = 0.0;
            output.originalRows = 0;
            output.yVariable = "";
         }


        private void computeYCapforOddTuples()
        {
            //MessageBox.Show("" + output.arrayXOriginalCopy[output.arrayXNoYRowNumbers.ElementAt(0),0]);
            //MessageBox.Show("" + output.arrayXOriginalCopy[output.arrayXNoYRowNumbers.ElementAt(1), 0]);
        }
    }
}
