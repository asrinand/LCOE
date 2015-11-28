using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace HAWT
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private void buttonLCOE_Click(object sender, EventArgs e)
        {
            if (checkBoxLoan.Checked == false)
            {
                if (textBoxRatedPower.Text != "" && textBoxRotorDiameter.Text != "" && textBoxAvgWindSpeed.Text != "" && textBoxCapitalCost.Text != "" && textBoxEquity.Text != "" && textBoxAnnualReturnEquity.Text != "" && textBoxOMCost.Text != "" && textBoxLoanInterest.Text != "" && textBoxLoanYears.Text != "")
                {
                    try
                    {
                        double temp, temp1, temp2, temp3, finaltemp;
                        double a, b, c, d, ee, f, g, h, i;
                        double crf, cf, AnnEnergyProd, AnnLoanPaym, OMCost, EquityReturn;

                        a = double.Parse(textBoxRatedPower.Text);
                        b = double.Parse(textBoxRotorDiameter.Text);
                        c = double.Parse(textBoxAvgWindSpeed.Text);
                        d = double.Parse(textBoxCapitalCost.Text);
                        ee = double.Parse(textBoxEquity.Text);
                        f = double.Parse(textBoxAnnualReturnEquity.Text);
                        g = double.Parse(textBoxOMCost.Text);
                        h = double.Parse(textBoxLoanInterest.Text);
                        i = double.Parse(textBoxLoanYears.Text);

                        //CRF - Capital Recovery Factor
                        temp = (1 + h);
                        temp1 = Math.Pow(temp, i);
                        temp2 = h * temp1;
                        temp3 = temp1 - 1;
                        crf = temp2 / temp3;
                        labelCRF.Text = crf.ToString("#.####");

                        //Capacity Factor
                        cf = 0.087 * c - (a / (b * b));
                        labelCapacityFactor.Text = cf.ToString("#.####");

                        //Annual Energy Production
                        AnnEnergyProd = a * 8760 * cf;
                        labelAnnualEnergyProd.Text = AnnEnergyProd.ToString("#.####");

                        //Annual Loan Payment
                        temp = (double)(1 + h);
                        temp1 = Math.Pow(temp, i);
                        temp2 = h * temp1;
                        temp3 = temp1 - 1;
                        crf = temp2 / temp3;
                        finaltemp = 1.00 - ee;
                        AnnLoanPaym = finaltemp * a * d * crf;
                        labelAnnualLoanPayment.Text = AnnLoanPaym.ToString("#.####");

                        //O&M Cost
                        OMCost = d * a * g;
                        labelOMCost.Text = OMCost.ToString("#.####");

                        //Equity Return
                        EquityReturn = d * ee * f * a;
                        labelEquityReturn.Text = EquityReturn.ToString("#.####");

                        //LCOE Final
                        double finalresult, a1, b1, c1, d1;
                        a1 = double.Parse(labelAnnualEnergyProd.Text);
                        b1 = double.Parse(labelAnnualLoanPayment.Text);
                        c1 = double.Parse(labelOMCost.Text);
                        d1 = double.Parse(labelEquityReturn.Text);
                        finalresult = (b1 + c1 + d1) / a1;
                        labelFINALVAL.Text = finalresult.ToString("$ #.##### /kWh");
                    }
                    catch(Exception E)
                    {
                        MessageBox.Show("Error : "
                              + E.Message.ToString());
                    }

                }
                else 
                {
                    MessageBox.Show("No text fields can be empty");
                }
            }
            else
            {
                
                if (textBoxRatedPower.Text != "" && textBoxRotorDiameter.Text != "" && textBoxAvgWindSpeed.Text != "" && textBoxCapitalCost.Text != "" && textBoxEquity.Text != "" && textBoxAnnualReturnEquity.Text != "" && textBoxOMCost.Text != "" && textBoxLoanInterest.Text == "" && textBoxLoanYears.Text == "")
                {
                    try
                    {
                        double a, b, c, d, ee, f, g;
                        double cf, AnnEnergyProd, OMCost, EquityReturn;

                        a = double.Parse(textBoxRatedPower.Text);
                        b = double.Parse(textBoxRotorDiameter.Text);
                        c = double.Parse(textBoxAvgWindSpeed.Text);
                        d = double.Parse(textBoxCapitalCost.Text);
                        ee = double.Parse(textBoxEquity.Text);
                        f = double.Parse(textBoxAnnualReturnEquity.Text);
                        g = double.Parse(textBoxOMCost.Text);

                        //Capacity Factor
                        cf = 0.087 * c - (a / (b * b));
                        labelCapacityFactor.Text = cf.ToString("#.####");

                        //Annual Energy Production
                        AnnEnergyProd = a * 8760 * cf;
                        labelAnnualEnergyProd.Text = AnnEnergyProd.ToString("#.####");

                        //O&M Cost
                        OMCost = d * a * g;
                        labelOMCost.Text = OMCost.ToString("#.####");

                        //Equity Return
                        EquityReturn = d * ee * f * a;
                        labelEquityReturn.Text = EquityReturn.ToString("#.####");

                        //LCOE Final
                        double finalresult, a1, c1, d1;
                        a1 = double.Parse(labelAnnualEnergyProd.Text);
                        c1 = double.Parse(labelOMCost.Text);
                        d1 = double.Parse(labelEquityReturn.Text);
                        finalresult = (c1 + d1) / a1;
                        labelFINALVAL.Text = finalresult.ToString("$ #.##### /kWh");
                    }
                    catch(Exception E1)
                    {
                        MessageBox.Show("Error : "
                              + E1.Message.ToString());
                    }

                }
                else
                {
                    MessageBox.Show("No text fields can be empty");
                }
            }
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void checkBoxLoan_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxLoan.Checked == true)
            {
                textBoxLoanInterest.Text = "";
                textBoxLoanInterest.Enabled = false;
                textBoxLoanYears.Text = "";
                textBoxLoanYears.Enabled = false;
            }
            else  
            {
                textBoxLoanInterest.Text = "";
                textBoxLoanInterest.Enabled = true;
                textBoxLoanYears.Text = "";
                textBoxLoanYears.Enabled = true;
            }
        }

        private void textBoxRatedPower_KeyPress(object sender, KeyPressEventArgs e)
        {
            //Allows only numbers and one decimal point
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != '\b' && e.KeyChar != '.')
            {
                e.Handled = true;
            }
            if (e.KeyChar == '.' && textBoxRatedPower.Text.Contains("."))
            {
                e.Handled = true;
            }
        }

        private void textBoxRotorDiameter_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != '\b' && e.KeyChar != '.')
            {
                e.Handled = true;
            }
            if (e.KeyChar == '.' && textBoxRotorDiameter.Text.Contains("."))
            {
                e.Handled = true;
            }
        }

        private void textBoxAvgWindSpeed_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != '\b' && e.KeyChar != '.')
            {
                e.Handled = true;
            }
            if (e.KeyChar == '.' && textBoxAvgWindSpeed.Text.Contains("."))
            {
                e.Handled = true;
            }
        }

        private void textBoxCapitalCost_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != '\b' && e.KeyChar != '.')
            {
                e.Handled = true;
            }
            if (e.KeyChar == '.' && textBoxCapitalCost.Text.Contains("."))
            {
                e.Handled = true;
            }
        }

        private void textBoxEquity_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != '\b' && e.KeyChar != '.')
            {
                e.Handled = true;
            }
            if (e.KeyChar == '.' && textBoxEquity.Text.Contains("."))
            {
                e.Handled = true;
            }
        }

        private void textBoxAnnualReturnEquity_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != '\b' && e.KeyChar != '.')
            {
                e.Handled = true;
            }
            if (e.KeyChar == '.' && textBoxAnnualReturnEquity.Text.Contains("."))
            {
                e.Handled = true;
            }
        }

        private void textBoxOMCost_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != '\b' && e.KeyChar != '.')
            {
                e.Handled = true;
            }
            if (e.KeyChar == '.' && textBoxOMCost.Text.Contains("."))
            {
                e.Handled = true;
            }
        }

        private void textBoxLoanInterest_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != '\b' && e.KeyChar != '.')
            {
                e.Handled = true;
            }
            if (e.KeyChar == '.' && textBoxLoanInterest.Text.Contains("."))
            {
                e.Handled = true;
            }
        }

        private void textBoxLoanYears_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != '\b' && e.KeyChar != '.')
            {
                e.Handled = true;
            }
            if (e.KeyChar == '.' && textBoxLoanYears.Text.Contains("."))
            {
                e.Handled = true;
            }
        }

        private void Form1_Load_1(object sender, EventArgs e)
        {
            this.MinimumSize = new Size(897, 435);
            this.MaximumSize = new Size(897, 435);
            this.MaximizeBox = false;
        }

        private void textBoxRatedPower_MouseHover(object sender, EventArgs e)
        {
            toolTip1.SetToolTip(this.textBoxRatedPower, "Enter Rated Power in Kilowatt");
        }

        private void textBoxRotorDiameter_MouseHover(object sender, EventArgs e)
        {
            toolTip1.SetToolTip(this.textBoxRotorDiameter, "Enter Rotor Diameter in metres");
        }

        private void textBoxAvgWindSpeed_MouseHover(object sender, EventArgs e)
        {
            toolTip1.SetToolTip(this.textBoxAvgWindSpeed, "Enter the Average Wind Speed in metre/second");
        }

        private void textBoxCapitalCost_MouseHover(object sender, EventArgs e)
        {
            toolTip1.SetToolTip(this.textBoxCapitalCost, "Enter the Capital Cost of the Wind Turbine");
        }

        private void textBoxEquity_MouseHover(object sender, EventArgs e)
        {
            toolTip1.SetToolTip(this.textBoxEquity, "Enter the Equity % as decimal number (eg : if 25% then enter 0.25)");
        }

        private void textBoxAnnualReturnEquity_MouseHover(object sender, EventArgs e)
        {
           toolTip1.SetToolTip(this.textBoxAnnualReturnEquity, "Enter the Annual Return on Equity % as decimal number (eg : if 25% then enter 0.25)");
        }

        private void textBoxOMCost_MouseHover(object sender, EventArgs e)
        {
            toolTip1.SetToolTip(this.textBoxOMCost, "Enter the Operation and Maintainence Cost % as decimal number (eg : if 3% then enter 0.03)");
        }

        private void textBoxLoanInterest_MouseHover(object sender, EventArgs e)
        {
            toolTip1.SetToolTip(this.textBoxLoanInterest, "Enter the Loan Interest % as decimal number (eg : if 25% then enter 0.25)");
        }

        private void textBoxLoanYears_MouseHover(object sender, EventArgs e)
        {
            toolTip1.SetToolTip(this.textBoxLoanYears, "Enter the Loan Period in years");
        }

        private void checkBoxLoan_MouseHover(object sender, EventArgs e)
        {
            toolTip1.SetToolTip(this.checkBoxLoan, "Select, if no loan is taken");
        }

        private void buttonLCOE_MouseHover(object sender, EventArgs e)
        {
            toolTip1.SetToolTip(this.buttonLCOE, "To Calculate the Levelized Cost of Electricity");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBoxRatedPower.Text = "";
            textBoxRotorDiameter.Text = "";
            textBoxAvgWindSpeed.Text = "";
            textBoxCapitalCost.Text = "";
            textBoxEquity.Text = "";
            textBoxAnnualReturnEquity.Text = "";
            textBoxOMCost.Text = "";
            textBoxLoanInterest.Text = "";
            textBoxLoanYears.Text = "";
            checkBoxLoan.Checked = false;
            labelFINALVAL.Text = "";
        }

    }
}
