using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MathNet.Numerics.LinearAlgebra;


namespace Terry_IN_BA_Regression
{
    public class OutputModel
    {
        public string[,] arrayX;
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
        public string confidenceLevel;
        public Boolean isLabelsCheckedInBasic;
    }
}
