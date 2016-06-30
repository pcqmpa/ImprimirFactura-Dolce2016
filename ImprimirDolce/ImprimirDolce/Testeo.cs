using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImprimirDolce
{

    
    public class Testeo
    {
        private PrintDocument _print = new PrintDocument();
        private int _numero;
        public Testeo()
        {
            _print.PrintPage += _print_PrintPage;
        }

        void _print_PrintPage(object sender, PrintPageEventArgs e)
        {
                e.Graphics.DrawString(_numero.ToString(),new Font("Arial",90,FontStyle.Bold),Brushes.Black,80,25);
        }

        public void Imprimir(int numero)
        {
            _numero = numero;
            _print.Print();
        }
    }
}
