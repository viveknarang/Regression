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
using System.Threading;

namespace Terry_IN_BA_Regression
{
    public partial class Form1 : Form
    {
        InputModel input = new InputModel();
        OutputModel output = new OutputModel();

        public Boolean textBox1Selected = false;
        public Boolean textBox2Selected = false;

        string[] xVariables;
        bool[] xVariableStates;

        public Form1()
        {
            InitializeComponent();

            checkBox41.Enabled = false;
            checkBox40.Enabled = false;
            checkBox39.Enabled = false;
            checkBox38.Enabled = false;
            checkBox37.Enabled = false;
            checkBox36.Enabled = false;
            checkBox35.Enabled = false;
            checkBox34.Enabled = false;
            checkBox33.Enabled = false;
            checkBox21.Enabled = false;

            checkBox41.Visible = false;
            checkBox40.Visible = false;
            checkBox39.Visible = false;
            checkBox38.Visible = false;
            checkBox37.Visible = false;
            checkBox36.Visible = false;
            checkBox35.Visible = false;
            checkBox34.Visible = false;
            checkBox33.Visible = false;
            checkBox21.Visible = false;

            textBox3.Enabled = Util.confidenceLevelInBasicChecked;
            textBox3.Text = Util.confidenceLevelDefaultValue;
        }

        private void initializeCheckboxesInVariablesSelect(string[] xVariables)
        {
            this.xVariables = xVariables;
            xVariableStates = new bool[xVariables.Length];

            for (int i = 1; i <= xVariables.Length; i++)
            {
                if (i == 1)
                {
                    if (xVariables[i - 1].Length <= 10)
                    {
                        checkBox41.Text = xVariables[i - 1];
                    } else
                    {
                        checkBox41.Text = xVariables[i - 1].Substring(0,9) + "...";
                    }
                    checkBox41.Enabled = true;
                    checkBox41.Checked = true;
                    checkBox41.Visible = true;
                    xVariableStates[i - 1] = true;
                }
                if (i == 2)
                {
                    if (xVariables[i - 1].Length <= 10)
                    {
                        checkBox40.Text = xVariables[i - 1];
                    }
                    else
                    {
                        checkBox40.Text = xVariables[i - 1].Substring(0, 9) + "...";
                    }
                    checkBox40.Enabled = true;
                    checkBox40.Checked = true;
                    checkBox40.Visible = true;
                    xVariableStates[i - 1] = true;
                }
                if (i == 3)
                {
                    if (xVariables[i - 1].Length <= 10)
                    {
                        checkBox39.Text = xVariables[i - 1];
                    }
                    else
                    {
                        checkBox39.Text = xVariables[i - 1].Substring(0, 9) + "...";
                    }
                    checkBox39.Enabled = true;
                    checkBox39.Checked = true;
                    checkBox39.Visible = true;
                    xVariableStates[i - 1] = true;
                }
                if (i == 4)
                {
                    if (xVariables[i - 1].Length <= 10)
                    {
                        checkBox38.Text = xVariables[i - 1];
                    }
                    else
                    {
                        checkBox38.Text = xVariables[i - 1].Substring(0, 9) + "...";
                    }
                    checkBox38.Enabled = true;
                    checkBox38.Checked = true;
                    checkBox38.Visible = true;
                    xVariableStates[i - 1] = true;
                }
                if (i == 5)
                {
                    if (xVariables[i - 1].Length <= 10)
                    {
                        checkBox37.Text = xVariables[i - 1];
                    }
                    else
                    {
                        checkBox37.Text = xVariables[i - 1].Substring(0, 9) + "...";
                    }
                    checkBox37.Enabled = true;
                    checkBox37.Checked = true;
                    checkBox37.Visible = true;
                    xVariableStates[i - 1] = true;
                }
                if (i == 6)
                {
                    if (xVariables[i - 1].Length <= 10)
                    {
                        checkBox36.Text = xVariables[i - 1];
                    }
                    else
                    {
                        checkBox36.Text = xVariables[i - 1].Substring(0, 9) + "...";
                    }
                    checkBox36.Enabled = true;
                    checkBox36.Checked = true;
                    checkBox36.Visible = true;
                    xVariableStates[i - 1] = true;
                }
                if (i == 7)
                {
                    if (xVariables[i - 1].Length <= 10)
                    {
                        checkBox35.Text = xVariables[i - 1];
                    }
                    else
                    {
                        checkBox35.Text = xVariables[i - 1].Substring(0, 9) + "...";
                    }
                    checkBox35.Enabled = true;
                    checkBox35.Checked = true;
                    checkBox35.Visible = true;
                    xVariableStates[i - 1] = true;
                }
                if (i == 8)
                {
                    if (xVariables[i - 1].Length <= 10)
                    {
                        checkBox34.Text = xVariables[i - 1];
                    }
                    else
                    {
                        checkBox34.Text = xVariables[i - 1].Substring(0, 9) + "...";
                    }
                    checkBox34.Enabled = true;
                    checkBox34.Checked = true;
                    checkBox34.Visible = true;
                    xVariableStates[i - 1] = true;
                }
                if (i == 9)
                {
                    if (xVariables[i - 1].Length <= 10)
                    {
                        checkBox33.Text = xVariables[i - 1];
                    }
                    else
                    {
                        checkBox33.Text = xVariables[i - 1].Substring(0, 9) + "...";
                    }
                    checkBox33.Enabled = true;
                    checkBox33.Checked = true;
                    checkBox33.Visible = true;
                    xVariableStates[i - 1] = true;
                }
                if (i == 10)
                {
                    if (xVariables[i - 1].Length <= 10)
                    {
                        checkBox21.Text = xVariables[i - 1];
                    }
                    else
                    {
                        checkBox21.Text = xVariables[i - 1].Substring(0, 9) + "...";
                    }
                    checkBox21.Enabled = true;
                    checkBox21.Checked = true;
                    checkBox21.Visible = true;
                    xVariableStates[i - 1] = true;
                }

            }

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
            InputModel input = Util.doSelectInputFromCurrentSheet();
            this.input = input;

            if (textBox1Selected == true)
            {
                textBox1.Text = input.cellnames;
                output.arrayY = input.array;
                output.yR = input.totalItems / input.columns.Count;
                output.yC = input.columns.Count;        
            }
            else if (textBox2Selected == true) 
            {
                textBox2.Text = input.cellnames;
                output.arrayX = input.array;
                output.xR = input.totalItems / input.columns.Count;
                output.xC = input.columns.Count;
                variableSelector();
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
            onCancelClick();
        }

        private void button4_Click(object sender, EventArgs e)
        {
        }

        private void button1_Click(object sender, EventArgs e)
        {
            onOkClick();
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

        private void variableSelector()
        {
            if (output.arrayX == null || output.arrayX.Length == 0)
            {
                MessageBox.Show("Please select independent variables first !", "Error !", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            string[] xVariable = new string[output.xC];

            double value;
            if (!double.TryParse(output.arrayX[0, 0], out value))
            {
                for (int i = 0; i < output.xC; i++)
                {
                    xVariable[i] = output.arrayX[0, i];
                }
            }
            else
            {
                for (int i = 0; i < output.xC; i++)
                {
                    xVariable[i] = "X" + (i + 1);
                }
            }

            initializeCheckboxesInVariablesSelect(xVariable);
        }

        public void InjectXVariableStates(LinkedList<bool> xVariableStates)
        {
            output.xVariableStates = xVariableStates;
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
        }

        private void initialize()
        {
            output.xVariableStates = new LinkedList<bool>(xVariableStates);
        }

        private void setInputStates()
        {
            output.noIntercept = checkBox2.Checked;
            output.isStandardizedCoefficientsEnabled = checkBox4.Checked;
            output.isOriginalEnabledInAdvancedOptions = checkBox6.Checked;
            output.isPredictedEnabledInAdvancedOptions = checkBox7.Checked;
            output.isConfidenceLimitsEnabledInAdvancedOptions = checkBox8.Checked;
            output.isResidualsEnabledInAdvancedOptions = checkBox9.Checked;
            output.isStandardizedResidualsEnabledInAdvancedOtions = checkBox10.Checked;
            output.isStudentizedResidualsEnabledInAdvancedOptions = checkBox11.Checked;
            output.isPRESSResidualsEnabledInAdvancedOptions = checkBox12.Checked;
            output.isRStudentEnabledInAdvancedOptions = checkBox13.Checked;
            output.isLeverageEnabledInAdvancedOptions = checkBox14.Checked;
            output.isCooksDEnabledInAdvancedOptions = checkBox15.Checked;
            output.isDFFITSEnabledInAdvancedOptions = checkBox16.Checked;
            output.isDFBETASEnabledInAdvancedOptions = checkBox17.Checked;
            output.isLabelsCheckedInBasic = checkBox1.Checked;
            output.noIntercept = checkBox2.Checked;
            output.confidenceLevel = textBox3.Text;
            output.isScatterPlotCheckedInPAndGSection = checkBox5.Checked;
            output.isScatterPlotCheckedInPAndGSection = checkBox20.Checked;
            output.isResidualsByPredictedCheckedInPAndGSection = checkBox22.Checked;
            output.isStandardizedResidualsByPredictedCheckedInPAndGSection = checkBox24.Checked;
            output.isResidualsByXVariablesCheckedInPAndGSection = checkBox23.Checked;
            output.isStandardizedResidualsByXVariablesCheckedInPAndGSection = checkBox25.Checked;
            output.isResidualsCheckedInPAndGSection = checkBox26.Checked;
            output.isYVariableCheckedInPAndGSection = checkBox28.Checked;
            output.isStandardizedResidualsCheckedInPAndGSection = checkBox27.Checked;
            output.isOtherCheckedInPAndGSection = checkBox29.Checked;
            output.isLeverageCheckedInPAndGSection = checkBox30.Checked;
            output.isDFFITSCheckedInPAndGSection = checkBox32.Checked;
            output.isCooksDCheckedInPAndGSection = checkBox31.Checked;
        }

        public void onOkClick()
        {
            setInputStates();
            initialize();

            Validator validator = new Validator(input, output);
            if (validator.validate())
            {
                return;
            }

            this.Hide();
            Form3 progress = new Terry_IN_BA_Regression.Form3();
            progress.Visible = true;
            
            ComputationCore core = new ComputationCore(input, output);
            OutputModel newOutput = core.getOutputModel();
            View view = new View(newOutput, input);
            view.createOutputOnASeparateSheet();
            core.clearCache();

            progress.Hide();
            this.Hide();
        }

        public void onCancelClick()
        {
            this.Visible = false;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            onOkClick();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            onCancelClick();
        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void toolStripStatusLabel1_Click(object sender, EventArgs e)
        {

        }

        public void updateStatus(String message)
        {
        }

        private void checkBox20_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox20.Checked)
            {
                checkBox5.Checked = true;
            }
            else
            {
                checkBox5.Checked = false;
            }
        }

        private void groupBox8_Enter(object sender, EventArgs e)
        {

        }

        private void checkBox30_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked)
            {
                checkBox20.Checked = true;
            } else
            {
                checkBox20.Checked = false;
            }
        }

        private void label21_Click(object sender, EventArgs e)
        {

        }

        private void label20_Click(object sender, EventArgs e)
        {

        }

        private void groupBox6_Enter(object sender, EventArgs e)
        {

        }

        private void checkBox23_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button14_Click(object sender, EventArgs e)
        {
            onOkClick();
        }

        private void button16_Click(object sender, EventArgs e)
        {
            onCancelClick();
        }

        private void button18_Click(object sender, EventArgs e)
        {
            onOkClick();
        }

        private void button20_Click(object sender, EventArgs e)
        {
            onCancelClick();
        }

        private void checkBox41_CheckedChanged(object sender, EventArgs e)
        {
            xVariableStates[0] = checkBox41.Checked;
        }

        private void checkBox40_CheckedChanged(object sender, EventArgs e)
        {
            xVariableStates[1] = checkBox40.Checked;
        }

        private void checkBox39_CheckedChanged(object sender, EventArgs e)
        {
            xVariableStates[2] = checkBox39.Checked;
        }

        private void checkBox38_CheckedChanged(object sender, EventArgs e)
        {
            xVariableStates[3] = checkBox38.Checked;
        }

        private void checkBox37_CheckedChanged(object sender, EventArgs e)
        {
            xVariableStates[4] = checkBox37.Checked;
        }

        private void checkBox36_CheckedChanged(object sender, EventArgs e)
        {
            xVariableStates[5] = checkBox36.Checked;
        }

        private void checkBox35_CheckedChanged(object sender, EventArgs e)
        {
            xVariableStates[6] = checkBox35.Checked;
        }

        private void checkBox34_CheckedChanged(object sender, EventArgs e)
        {
            xVariableStates[7] = checkBox34.Checked;
        }

        private void checkBox33_CheckedChanged(object sender, EventArgs e)
        {
            xVariableStates[8] = checkBox33.Checked;
        }

        private void checkBox21_CheckedChanged(object sender, EventArgs e)
        {
            xVariableStates[9] = checkBox21.Checked;
        }

        private void checkBox42_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox42.Checked == true)
            {
                checkBox6.Checked = true;
                checkBox7.Checked = true;
                checkBox8.Checked = true;
                checkBox9.Checked = true;
                checkBox10.Checked = true;
                checkBox11.Checked = true;
                checkBox12.Checked = true;
                checkBox13.Checked = true;
                checkBox14.Checked = true;
                checkBox15.Checked = true;
                checkBox16.Checked = true;
                checkBox17.Checked = true;
                checkBox18.Checked = true;
                checkBox19.Checked = true;
            } else
            {
                checkBox6.Checked = false;
                checkBox7.Checked = false;
                checkBox8.Checked = false;
                checkBox9.Checked = false;
                checkBox10.Checked = false;
                checkBox11.Checked = false;
                checkBox12.Checked = false;
                checkBox13.Checked = false;
                checkBox14.Checked = false;
                checkBox15.Checked = false;
                checkBox16.Checked = false;
                checkBox17.Checked = false;
                checkBox18.Checked = false;
                checkBox19.Checked = false;

            }
            setInputStates();
        }

        private void checkBox43_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox43.Checked == true)
            {
                checkBox20.Checked = true;
                checkBox22.Checked = true;
                checkBox23.Checked = true;
                checkBox24.Checked = true;
                checkBox25.Checked = true;
                checkBox26.Checked = true;
                checkBox27.Checked = true;
                checkBox28.Checked = true;
                checkBox29.Checked = true;
                checkBox30.Checked = true;
                checkBox31.Checked = true;
                checkBox32.Checked = true;
            } else
            {
                checkBox20.Checked = false;
                checkBox22.Checked = false;
                checkBox23.Checked = false;
                checkBox24.Checked = false;
                checkBox25.Checked = false;
                checkBox26.Checked = false;
                checkBox27.Checked = false;
                checkBox28.Checked = false;
                checkBox29.Checked = false;
                checkBox30.Checked = false;
                checkBox31.Checked = false;
                checkBox32.Checked = false;
            }
            setInputStates();
        }
    }
}