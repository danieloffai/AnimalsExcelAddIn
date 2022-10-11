using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace ExcelAddInZoo
{
    public partial class MessageBoxAnimalNumber : Form
    {
        public int AnimalNumber { get; private set; }

        public MessageBoxAnimalNumber()
        {
            InitializeComponent();
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (CheckNumber())
            {
                this.AnimalNumber = int.Parse(tboxAnimalNumber.Text);
                this.Close();
            }
        }

        private bool CheckNumber()
        {
            if (!int.TryParse(tboxAnimalNumber.Text, out int n))
            {
                MessageBox.Show("Only whole numbers are accpeted!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (n < 1 || n > 10)
            {
                MessageBox.Show("Only numbers between 1 and 10 are allowed!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            return true;
        }
    }
}
