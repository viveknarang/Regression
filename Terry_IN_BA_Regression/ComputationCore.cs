﻿using MathNet.Numerics.LinearAlgebra;
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
            computeLLMeanAndULMean();
        }

        private void init()
        {

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


        }

        private void adjustRowsForNullEntries()
        {

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
                for (j = 0; j < output.xC; j++)
                {
                    if (output.arrayX[i, j] == "")
                    {
                        removeRowSet.Add(i);
                    }
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
                    y = 0;
                    for (j = 0; j < output.yC; j++)
                    {
                        output.arrayYFilter[x, y] = output.arrayY[i, j];
                        y++;
                    }
                    x++;
                }
            }
            output.arrayY = output.arrayYFilter;
            output.xR = output.yR = output.arrayYFilter.Length;
            output.n = output.arrayYFilter.Length;
        }

        private void computeBetaCoefficients()
        {
            output.betaCoefficients = ((output.arrayXConverted.Transpose().Multiply(output.arrayXConverted)).Inverse()).Multiply(output.arrayXConverted.Transpose()).Multiply(output.arrayYConverted);
        }

        private void computeK()
        {
            output.k = output.xVariables.Count;
        }

        private void computeDfe()
        {
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
            output.yCap = output.arrayXConverted.Multiply(output.betaCoefficients);
        }

        private void computeYBar()
        {
            double ySum = 0;
            for (int i = 0; i < output.arrayYConverted.RowCount; i++)
            {
                ySum += output.arrayYConverted.At(i, 0);
            }
            output.yBar = ySum / output.arrayYConverted.RowCount;
        }

        private void computeSSR()
        {
            for (int i = 0; i < output.yCap.RowCount; i++)
            {
                output.SSR += (output.yCap.At(i, 0) - output.yBar) * (output.yCap.At(i, 0) - output.yBar);
            }
        }

        private void computeSSE()
        {
            for (int i = 0; i < output.yCap.RowCount; i++)
            {
                output.SSE += (output.arrayYConverted.At(i, 0) - output.yCap.At(i, 0)) * (output.arrayYConverted.At(i, 0) - output.yCap.At(i, 0));
            }
        }

        private void computeSST()
        {
            for (int i = 0; i < output.yCap.RowCount; i++)
            {
                output.SST += (output.arrayYConverted.At(i, 0) - output.yBar) * (output.arrayYConverted.At(i, 0) - output.yBar);
            }
        }

        private void computeMSR()
        {
            output.MSR = output.SSR / output.k;
        }

        private void computeMSE()
        {
            output.MSE = output.SSE / output.dfe;
        }

        private void computeFStatistic()
        {
            output.F = output.MSR / output.MSE;
        }

        private void computePValue()
        {
            output.P = MathNet.Numerics.ExcelFunctions.FDist(output.F, (int)output.k, (int)(output.dfe));
        }

        private void computeRSquare()
        {
            output.RSquare = output.SSR / output.SST;
        }

        private void computeR()
        {
            output.R = Math.Sqrt(output.RSquare);
        }

        private void computeRSQAdjusted()
        {
            output.RSqAdj = (((output.n - 1) * output.RSquare) - output.k) / (output.dfe);
        }

        private void computeSE()
        {
            output.SE = output.MSE * output.arrayXConverted.Transpose().Multiply(output.arrayXConverted).Inverse();

            for (int i = 0; i < output.SE.RowCount; i++)
            {
                output.SE[i, 0] = Math.Sqrt(output.SE.At(i, i));
            }
        }

        private void initializeMatrices()
        {
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

            p = new double[1, output.xVariables.Count];

            for (int i = 0; i < output.xVariables.Count; i++)
            {
                p[0, i] = 1;
            }

            output.SDx = DenseMatrix.OfArray(p);
        }

        private void computeStandardizedCoefficients()
        {
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
            for (int i = 0; i < output.SE.RowCount; i++)
            {
                output.tStat[i, 0] = output.betaCoefficients.At(i, 0) / output.SE.At(i, 0);
            }
        }

        private void computePValueForBetaCoefficients()
        {
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
            output.tStar = Math.Abs(MathNet.Numerics.Distributions.StudentT.InvCDF(0d, 1d, (int)output.dfe, (1 - (double.Parse(output.confidenceLevel) / 100)) / 2));
        }

        private void computeLLAndUL()
        {
            for (int i = 0; i < output.LL.RowCount; i++)
            {
                output.LL[i, 0] = output.betaCoefficients[i, 0] - output.tStar * output.SE[i, 0];
                output.UL[i, 0] = output.betaCoefficients[i, 0] + output.tStar * output.SE[i, 0];
            }
        }

        private void computeVIF()
        {

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

        private void computeLLMeanAndULMean()
        {
            for (int i = 0; i < output.arrayXConverted.RowCount; i++)
            {
                Matrix<double> m1 = output.arrayXConverted.Row(i).ToRowMatrix();
                Matrix<double> m2 = output.arrayXConverted.Transpose();
                Matrix<double> m3 = output.arrayXConverted;
                Matrix<double> m4 = output.arrayXConverted.Row(i).ToRowMatrix().Transpose();

                output.LLMean[i, 0] = output.yCap[i, 0] - output.tStar * Math.Sqrt(output.MSE * (m1.Multiply(((m2.Multiply(m3)).Inverse())).Multiply(m4)[0, 0]));
                output.ULMean[i, 0] = output.yCap[i, 0] + output.tStar * Math.Sqrt(output.MSE * (m1.Multiply(((m2.Multiply(m3)).Inverse())).Multiply(m4)[0, 0]));
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
}
}