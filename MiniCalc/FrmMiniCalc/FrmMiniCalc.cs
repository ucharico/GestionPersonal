using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MiniCalc;
namespace FrmMiniCalc
{
    public partial class FrmMiniCalc : Form
    {
        private CalcOperation op = CalcOperation.eAdd;
        public FrmMiniCalc()
        {
            InitializeComponent();
        }

        private void EqualsButton_Click(object sender, EventArgs e)
        {
            try
            {
                Calculator calc = new Calculator();
                Result.Text = calc.Calculate((int)NumberA.Value, op, (int)NumberB.Value).ToString();
            }
            catch (ResultOutOfRangeException)
            {
                Result.Text = "Out of Range";
            }
            catch (NegativeParameterException)
            {
                Result.Text = "Negatives not allowed";
            }
        }

        private void AddRButton_CheckedChanged(object sender, EventArgs e)
        {
            if (AddRButton.Checked)
            {
                op = CalcOperation.eAdd;
            }
        }

        private void SubtractRButton_CheckedChanged(object sender, EventArgs e)
        {
            if (SubtractRButton.Checked)
            {
                op = CalcOperation.eSubtract;
            }
        }

        private void FrmMiniCalc_Load(object sender, EventArgs e)
        {

        }

        private void CmdSalir_Click(object sender, EventArgs e)
        {
            this.Close();
            this.Dispose();
        }
    }
}
