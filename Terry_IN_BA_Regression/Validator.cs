using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Terry_IN_BA_Regression
{
    public class Validator
    {
        InputModel input;
        OutputModel output;
        public Validator(InputModel input, OutputModel output)
        {
            this.input = input;
            this.output = output;
        }

        public bool validate()
        {
            if (output.arrayX == null || output.arrayY == null)
            {
                Util.showErrorMessage("No data selected !", "Error !");
                return true;
            }

            if (output.yC > 1)
            {
                Util.showErrorMessage("Dependent variable has to be a single column with one or more rows !", "Error !");
                return true;
            }

            if (output.yR != output.xR)
            {
                Util.showErrorMessage("The rows of dependent and independent variables must match !", "Error !");
                return true;
            }

            if (output.isLabelsCheckedInBasic)
            {
                double number;
                bool isXNumeric = double.TryParse(output.arrayX[0, 0], out number);
                bool isYNumeric = double.TryParse(output.arrayY[0, 0], out number);
                if (isXNumeric || isYNumeric)
                {
                    Util.showErrorMessage("It looks like you expect labels but the first row of input does not look like having labels !", "Error !");
                    return true;
                }

                for (int i = 1; i < output.xR; i++)
                {
                    for (int j = 0; j < output.xC; j++)
                    {
                        if (output.arrayX[i, j] != "" && !double.TryParse(output.arrayX[i, j], out number))
                        {
                            Util.showErrorMessage("It appears that some of the independent variable input data are non-numeric - Specify only numeric input !", "Error !");
                            return true;
                        }
                    }
                }

                /*for (int i = 1; i < output.yR; i++)
                {
                    for (int j = 0; j < output.yC; j++)
                    {

                        if (!double.TryParse(output.arrayY[i, j], out number))
                        {
                            Util.showErrorMessage("It appears that some of the dependent variable input data are non-numeric - Specify only numeric input !", "Error !");
                            return true;
                        }

                    }

                }*/

            }
            else
            {
                double number;
                bool isXNumeric = double.TryParse(output.arrayX[0, 0], out number);
                bool isYNumeric = double.TryParse(output.arrayY[0, 0], out number);

                if (!isXNumeric || !isYNumeric)
                {
                    Util.showErrorMessage("It looks like you are not expecting labels but the first row of input does look like having labels !", "Error !");
                    return true;
                }

                for (int i = 0; i < output.xR; i++)
                {
                    for (int j = 0; j < output.xC; j++)
                    {

                        if (output.arrayX[i, j] != "" && !double.TryParse(output.arrayX[i, j], out number))
                        {
                            Util.showErrorMessage("It appears that some of the independent variable input data are non-numeric - Specify only numeric input !", "Error !");
                            return true;
                        }

                    }
                }

                /*for (int i = 0; i < output.yR; i++)
                {
                    for (int j = 0; j < output.yC; j++)
                    {

                        if (!double.TryParse(output.arrayY[i, j], out number))
                        {
                            Util.showErrorMessage("It appears that some of the dependent variable input data are non-numeric - Specify only numeric input !", "Error !");
                            return true;
                        }

                    }

                }*/

            }

            double num;

            if (!double.TryParse(output.confidenceLevel, out num))
            {
                Util.showErrorMessage("Confidence level does not look like a number - only numbers please !", "Error !");
                return true;
            }
            else
            {
                if (num < 0 || num > 99.9999)
                {
                    Util.showErrorMessage("Confidence level should be above 0 and less than 100 !", "Error !");
                    return true;
                }
            }

            return false;
        }
    }
}
