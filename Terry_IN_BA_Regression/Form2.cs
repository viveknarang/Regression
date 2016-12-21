using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Terry_IN_BA_Regression
{
    public partial class Form2 : Form
    {
        string[] xVariables;
        bool[] xVariableStates;

        public Form2(string[] xVariables)
        {
            InitializeComponent();
            checkBox1.Enabled = false;
            checkBox2.Enabled = false;
            checkBox3.Enabled = false;
            checkBox4.Enabled = false;
            checkBox5.Enabled = false;
            checkBox6.Enabled = false;
            checkBox7.Enabled = false;
            checkBox8.Enabled = false;
            checkBox9.Enabled = false;
            checkBox10.Enabled = false;

            checkBox1.Visible = false;
            checkBox2.Visible = false;
            checkBox3.Visible = false;
            checkBox4.Visible = false;
            checkBox5.Visible = false;
            checkBox6.Visible = false;
            checkBox7.Visible = false;
            checkBox8.Visible = false;
            checkBox9.Visible = false;
            checkBox10.Visible = false;

            this.xVariables = xVariables;
            xVariableStates = new bool[xVariables.Length];

            for (int i = 1; i <= xVariables.Length; i++)
            {
                if (i == 1)
                {
                    checkBox1.Text = xVariables[i-1];
                    checkBox1.Enabled = true;
                    checkBox1.Checked = true;
                    checkBox1.Visible = true;
                    xVariableStates[i - 1] = true;
                }
                if (i == 2)
                {
                    checkBox2.Text = xVariables[i-1];
                    checkBox2.Enabled = true;
                    checkBox2.Checked = true;
                    checkBox2.Visible = true;
                    xVariableStates[i - 1] = true;
                }
                if (i == 3)
                {
                    checkBox3.Text = xVariables[i-1];
                    checkBox3.Enabled = true;
                    checkBox3.Checked = true;
                    checkBox3.Visible = true;
                    xVariableStates[i - 1] = true;
                }
                if (i == 4)
                {
                    checkBox4.Text = xVariables[i-1];
                    checkBox4.Enabled = true;
                    checkBox4.Checked = true;
                    checkBox4.Visible = true;
                    xVariableStates[i - 1] = true;
                }
                if (i == 5)
                {
                    checkBox5.Text = xVariables[i-1];
                    checkBox5.Enabled = true;
                    checkBox5.Checked = true;
                    checkBox5.Visible = true;
                    xVariableStates[i - 1] = true;
                }
                if (i == 6)
                {
                    checkBox6.Text = xVariables[i-1];
                    checkBox6.Enabled = true;
                    checkBox6.Checked = true;
                    checkBox6.Visible = true;
                    xVariableStates[i - 1] = true;
                }
                if (i == 7)
                {
                    checkBox7.Text = xVariables[i-1];
                    checkBox7.Enabled = true;
                    checkBox7.Checked = true;
                    checkBox7.Visible = true;
                    xVariableStates[i - 1] = true;
                }
                if (i == 8)
                {
                    checkBox8.Text = xVariables[i-1];
                    checkBox8.Enabled = true;
                    checkBox8.Checked = true;
                    checkBox8.Visible = true;
                    xVariableStates[i - 1] = true;
                }
                if (i == 9)
                {
                    checkBox9.Text = xVariables[i-1];
                    checkBox9.Enabled = true;
                    checkBox9.Checked = true;
                    checkBox9.Visible = true;
                    xVariableStates[i - 1] = true;
                }
                if (i == 10)
                {
                    checkBox10.Text = xVariables[i - 1];
                    checkBox10.Enabled = true;
                    checkBox10.Checked = true;
                    checkBox10.Visible = true;
                    xVariableStates[i - 1] = true;
                }

            }
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            LinkedList<bool> xvS = new LinkedList<bool>();
            for(int i = 0; i < xVariableStates.Length; i++)
            {
                xvS.AddLast(xVariableStates[i]);
            }

            Terry_IN_BA_Regression.ThisAddIn.form.InjectXVariableStates(xvS);
            this.Hide();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            xVariableStates[0] = checkBox1.Checked;
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            xVariableStates[1] = checkBox2.Checked;
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            xVariableStates[2] = checkBox3.Checked;
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            xVariableStates[3] = checkBox4.Checked;
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            xVariableStates[4] = checkBox5.Checked;
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            xVariableStates[5] = checkBox6.Checked;
        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            xVariableStates[6] = checkBox7.Checked;
        }

        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {
            xVariableStates[7] = checkBox8.Checked;
        }

        private void checkBox9_CheckedChanged(object sender, EventArgs e)
        {
            xVariableStates[8] = checkBox9.Checked;
        }

        private void checkBox10_CheckedChanged(object sender, EventArgs e)
        {
            xVariableStates[9] = checkBox10.Checked;
        }
    }
}
