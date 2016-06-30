using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ImprimirDolce;

namespace test
{
    class Program
    {
        static void Main(string[] args)
        {

            //var impresion = new ImprimirFacturaDolce();
            //var impresion = new ImprimirFacturaDolce();

            var impresion = new ImprimirFacturaDolce ();
            var mensaje = "";
            bool _logrespuesta;
            string xarchivo = string.Empty;

            impresion.ImprimirDatos(1384050, "", out mensaje, out _logrespuesta);

            //impresion.FacturaPdf(1240949, "D:\\",out mensaje, out _logrespuesta, out xarchivo);


            //impresion.ImprimirDatos(620087,"",out mensaje,out _logrespuesta);
            //var impresion = new Testeo();
            //int numero = 0;
            //for (int i = 0; i < 101; i++)
            //{
            //    numero = i + 1;
            //    impresion.Imprimir(numero);
            //}



            //if (_logrespuesta)
            //{
            //  Console.WriteLine("Listo");  
            //}

            Console.WriteLine("Ok");
            Console.ReadKey();
        }
    }
}
