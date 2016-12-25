using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace Terry_IN_BA_Regression
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            if (Object.ReferenceEquals(ThisAddIn.form,null))
            {
                ThisAddIn.form.Visible = true;
                ThisAddIn.form.TopMost = true;
            }
            else
            {
                Form1 form = new Form1();
                ThisAddIn.form = form;
                ThisAddIn.form.Visible = true;
                ThisAddIn.form.TopMost = true;
            }
        }
    }
}
