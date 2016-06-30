using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PruebaImpresion
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var vc = new ImprimirDolce.ImprimirFacturaDolce();

            string strMensaje;

            var res =vc.ImprimirDatos(511436, out strMensaje);
            if (res)
            {
                MessageBox.Show("OK");
            }
            else
            {
                MessageBox.Show(strMensaje);
            }

        }
    }
}
