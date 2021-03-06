﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MathNet.Numerics.LinearAlgebra;


namespace Terry_IN_BA_Regression
{
    public class OutputModel
    {
        public string[,] arrayXOriginalCopy;
        public string[,] arrayYOriginalCopy;
        public string[,] arrayX;
        public HashSet<int> arrayXNoYRowNumbers = new HashSet<int>();
        public string[,] arrayY;
        public string[,] arrayXFilter;
        public string[,] arrayYFilter;
        public LinkedList<string> xVariables = new LinkedList<string>();
        public string yVariable;
        public LinkedList<bool> xVariableStates = new LinkedList<bool>();
        public double n;
        public double k;
        public int yR = 0;
        public int yC = 0;
        public int xR = 0;
        public int xC = 0;
        public Matrix<double> yCap;
        public double SSR;
        public double yBar;
        public double SSE;
        public double SST;
        public double MSR;
        public double MSE;
        public double F;
        public double P;
        public double RSquare;
        public double R;
        public double RSqAdj;
        public Matrix<double> betaCoefficients;
        public Matrix<double> SE;
        public Matrix<double> tStat;
        public Matrix<double> pValue;
        public Matrix<double> LL;
        public Matrix<double> UL;
        public Matrix<double> SDx;
        public Matrix<double> standardizedCoefficients;
        public Matrix<double> LLMean;
        public Matrix<double> ULMean;
        public Matrix<double> LLPred;
        public Matrix<double> ULPred;
        public Matrix<double> residuals;
        public Matrix<double> standardizedResiduals;
        public Matrix<double> studentizedResiduals;
        public Matrix<double> PRESSResiduals;
        public Matrix<double> RStudentResiduals;
        public Matrix<double> Leverage;
        public Matrix<double> CooksD;
        public Matrix<double> DFFITS;
        public Matrix<double> DFBETAS;
        public Matrix<double> arrayXNoY;
        public Matrix<double> arrayXNoYComputedY;
        public Matrix<double> arrayXNoYLLMean;
        public Matrix<double> arrayXNoYULMean;
        public Matrix<double> arrayXNoYLLPred;
        public Matrix<double> arrayXNoYULPred;
        public Matrix<double> cumulativeProportionForResiduals;
        public Matrix<double> standardNormalQuantileForResiduals;
        public Matrix<double> cumulativeProportionForStandardizedResiduals;
        public Matrix<double> standardNormalQuantileForStandardizedResiduals;
        public Matrix<double> cumulativeProportionForDependentVariable;
        public Matrix<double> standardNormalQuantileForDependentVariable;
        public Matrix<double> observationNumber;
        public int sampleSize; 
        public double SDy;
        public double dfe;
        public double tStar;
        public Matrix<double> VIFMatrix;
        public int originalRows;
        public Matrix<double> arrayYConverted;
        public Matrix<double> arrayXConverted;
        public Boolean noIntercept;
        public Boolean isStandardizedCoefficientsEnabled;
        public Boolean isOriginalEnabledInAdvancedOptions;
        public Boolean isPredictedEnabledInAdvancedOptions;
        public Boolean isConfidenceLimitsEnabledInAdvancedOptions;
        public Boolean isResidualsEnabledInAdvancedOptions;
        public Boolean isStandardizedResidualsEnabledInAdvancedOtions;
        public Boolean isStudentizedResidualsEnabledInAdvancedOptions;
        public Boolean isPRESSResidualsEnabledInAdvancedOptions;
        public Boolean isRStudentEnabledInAdvancedOptions;
        public Boolean isLeverageEnabledInAdvancedOptions;
        public Boolean isCooksDEnabledInAdvancedOptions;
        public Boolean isDFFITSEnabledInAdvancedOptions;
        public Boolean isDFBETASEnabledInAdvancedOptions;
        public string confidenceLevel;
        public Boolean isLabelsCheckedInBasic;
        public Boolean isScatterPlotCheckedInPAndGSection;
        public Boolean isResidualsByPredictedCheckedInPAndGSection;
        public Boolean isStandardizedResidualsByPredictedCheckedInPAndGSection;
        public Boolean isResidualsByXVariablesCheckedInPAndGSection;
        public Boolean isStandardizedResidualsByXVariablesCheckedInPAndGSection;
        public Boolean isResidualsCheckedInPAndGSection;
        public Boolean isYVariableCheckedInPAndGSection;
        public Boolean isStandardizedResidualsCheckedInPAndGSection;
        public Boolean isOtherCheckedInPAndGSection;
        public Boolean isLeverageCheckedInPAndGSection;
        public Boolean isDFFITSCheckedInPAndGSection;
        public Boolean isCooksDCheckedInPAndGSection;
    }
}
