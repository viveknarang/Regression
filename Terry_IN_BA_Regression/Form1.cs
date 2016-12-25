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
        InputModel input = new InputModel();
        OutputModel output = new OutputModel();

        public Boolean textBox1Selected = false;
        public Boolean textBox2Selected = false;

        public Form1()
        {
            InitializeComponent();
            textBox3.Enabled = Util.confidenceLevelInBasicChecked;
            textBox3.Text = Util.confidenceLevelDefaultValue;
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

        private void button5_Click(object sender, EventArgs e)
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
                    xVariable[i] = "X" + (i+1);
                }
            }

            Form2 form = new Form2(xVariable);
            form.Visible = true;
            form.TopMost = true;
        }

        public void InjectXVariableStates(LinkedList<bool> xVariableStates)
        {
            output.xVariableStates = xVariableStates;
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
        }

        public void onOkClick()
        {
            output.noIntercept = checkBox2.Checked;
            output.isStandardizedCoefficientsEnabled = checkBox4.Checked;
            output.isOriginalEnabledInAdvancedOptions = checkBox6.Checked;
            output.isPredictedEnabledInAdvancedOptions = checkBox7.Checked;
            output.isConfidenceLimitsEnabledInAdvancedOptions = checkBox8.Checked;
            output.isResidualsEnabledInAdvancedOptions = checkBox9.Checked;
            output.isLabelsCheckedInBasic = checkBox1.Checked;
            output.noIntercept = checkBox2.Checked;
            output.confidenceLevel = textBox3.Text;

            Validator validator = new Validator(input,output);
            if (validator.validate())
            {
                return;
            }

            ComputationCore core = new ComputationCore(input, output);
            OutputModel newOutput = core.getOutputModel();
            View view = new View(newOutput);
            view.createOutputOnASeparateSheet();
            core.clearCache();
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
    }
}