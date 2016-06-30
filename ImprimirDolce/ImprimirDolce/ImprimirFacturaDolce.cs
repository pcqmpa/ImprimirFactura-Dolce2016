using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using BarcodeLib.Barcode;

namespace ImprimirDolce
{

    public interface IFacturaDolce
    {
        void ImprimirDatos(Int64 numPedido,string strImpresora ,out string strMensaje, out bool logrespuesta);
    }

    [Export(typeof(IFacturaDolce))]
    public class ImprimirFacturaDolce:IFacturaDolce
    {
        private int paginaActual = 0;
        private PrintDocument _print = new PrintDocument();
        //private PrintDocument _post=new PrintDocument();
        private DataSet _dts;
        private DataHelper _dataHelper;
        private DataTable _detalle;
        private int _totalPaginas=0;
        private int _numPagina=0;

        [ImportingConstructor]
        public ImprimirFacturaDolce()
        {
            _dataHelper = new DataHelper(new SqlConnectionStringBuilder(Properties.Settings.Default.conexion_dbDolce));

            _print.PrintPage += _print_PrintPage;
            //_post.PrintPage += _post_PrintPage;
        }

        void _post_PrintPage(object sender, PrintPageEventArgs e)
        {

            var dtEncabezado = _dts.Tables[0];
            var dtInformativo = _dts.Tables[4];
            var dtPremios = _dts.Tables[3];

            int fontsize = 8;
            var xbrushes = Brushes.Black;
            var fueteTitulo = new Font("Tahoma", 9, FontStyle.Bold);
            var fuenteNormal = new Font("Tahoma", fontsize);
            var fuentepequeña = new Font("Tahoma", 5, FontStyle.Bold);
            var fuenteDetalle = new Font("Tahoma", 7, FontStyle.Bold);
            var fuentecolilla = new Font("Tahoma", 8, FontStyle.Bold);
            var fuentelistaDetalle = new Font("Arial", 7);
            var fuenteBancos = new Font("Arial", 5);



            //string rutaImagen = Properties.Settings.Default.Carpeta_Imagen.Trim();
            //Image logo = Image.FromFile(string.Format("{0}Logo_Dolce.jpg", rutaImagen));

            //Paginacion 
           
            //fin Paginacion

            //e.Graphics.DrawImage(logo, 270, 5);
            e.Graphics.DrawString(dtEncabezado.Rows[0]["nombreempresa"].ToString().Trim(), fueteTitulo, xbrushes, 350, 10);
            e.Graphics.DrawString(string.Format("Nit:{0}", dtEncabezado.Rows[0]["nitempresa"]).ToString().Trim(), fueteTitulo, xbrushes, 350, 25);
            e.Graphics.DrawString(dtEncabezado.Rows[0]["ventadirecta"].ToString().Trim(), fueteTitulo, xbrushes, 350, 40);
            e.Graphics.DrawString(string.Format("FACTURA DE VENTA No {0}-{1}", dtEncabezado.Rows[0]["Prefijo"].ToString().Trim(), dtEncabezado.Rows[0]["numFactura"].ToString().Trim()), fueteTitulo, xbrushes, 500, 10);
            e.Graphics.DrawString(string.Format("PEDIDO No {0}", dtEncabezado.Rows[0]["numPedido"].ToString().Trim()), fueteTitulo, xbrushes, 500, 25);
            e.Graphics.DrawString("CAMPAÑA", fueteTitulo, xbrushes, 730, 10);
            e.Graphics.DrawString(dtEncabezado.Rows[0]["StrCampaña"].ToString().Trim(), fueteTitulo, xbrushes, 730, 25);

            //Caja del Numero de pedido zona y total de prendas

            e.Graphics.FillRectangle(Brushes.LightGray, 10, 10, 250, 40);

            e.Graphics.DrawString("PEDIDOS No", fuenteBancos, xbrushes, 15, 10);
            e.Graphics.DrawString("ZONA", fuenteBancos, xbrushes, 110, 10);
            e.Graphics.DrawString("TOTAL PRENDAS", fuenteBancos, xbrushes, 170, 10);

            var listaTotalInformacion = dtInformativo.Select("strtipo=7");

            e.Graphics.DrawString(listaTotalInformacion[0]["strObservacion"].ToString().Trim(), new Font("Arial", 20, FontStyle.Bold), xbrushes, 15, 20);
            e.Graphics.DrawString(listaTotalInformacion[0]["strCodigo"].ToString().Trim(), new Font("Arial", 20, FontStyle.Bold), xbrushes, 90, 20);
            e.Graphics.DrawString(listaTotalInformacion[0]["intCantidad"].ToString().Trim(), new Font("Arial", 20, FontStyle.Bold), xbrushes, 180, 20);

            //Fin caja 

            //Detalles de la Pagina posteriror

            var pen = new Pen(Color.LightGray) { DashStyle = DashStyle.Dot };


            e.Graphics.FillRectangle(Brushes.LightGray, 10, 55, 820, 15);
            e.Graphics.DrawString("INFORMATIVO DOLCE", fueteTitulo, xbrushes, 15, 55);


            //Lineas Divisorias
            //Premios Entragados

            e.Graphics.DrawString("PREMIOS ENTREGADOS", fuentecolilla, xbrushes, 15, 70);

            int conteo = 75;

            for (int i = 0; i < dtPremios.Rows.Count; i++)
            {
                conteo += 10;
                e.Graphics.DrawString(string.Format("{0} CANTIDAD {1}", dtPremios.Rows[i]["strCodigo"].ToString().Trim(), dtPremios.Rows[i]["intCantidad"]), fuentelistaDetalle, xbrushes, 15, conteo);
            }

            e.Graphics.DrawLine(pen, 500, 60, 500, 830);

            //Puntos Acumulados

            var listaPuntos = dtInformativo.Select("strTipo=2");

            e.Graphics.DrawString("PUNTOS ACUMULADOS", fuentecolilla, xbrushes, 15, 255);

            conteo = 260;

            for (int i = 0; i < listaPuntos.Count(); i++)
            {
                if (i == 0)
                {
                    conteo += 10;
                    e.Graphics.DrawString(listaPuntos[i]["strObservacion"].ToString().Trim(), new Font("Arial",14,FontStyle.Bold), xbrushes, 15, conteo);
                }
                else
                {
                    conteo += 10;
                    e.Graphics.DrawString(listaPuntos[i]["strObservacion"].ToString().Trim(), fuenteBancos, xbrushes, 15, conteo);    
                }
                
            }
            e.Graphics.DrawLine(pen, 10, 250, 500, 250);


            //Recordatorio Otros

            e.Graphics.DrawString("RECORDATORIO OTROS", fuentecolilla, xbrushes, 15, 505);


            e.Graphics.DrawLine(pen, 10, 500, 500, 500);

            //Productos Agotados
            e.Graphics.FillRectangle(Brushes.LightGray, 500, 85, 330, 15);

            e.Graphics.DrawLine(pen, 600, 90, 600, 525);
            e.Graphics.DrawLine(pen, 780, 90, 780, 525);
            e.Graphics.DrawLine(pen, 500, 525, 830, 525);

            //Cambios Surtidos
            e.Graphics.FillRectangle(Brushes.LightGray, 500, 540, 330, 15);
            e.Graphics.DrawLine(pen, 600, 550, 600, 700);
            e.Graphics.DrawLine(pen, 780, 550, 780, 700);
            e.Graphics.DrawLine(pen, 500, 700, 830, 700);

            //Cambios Agotados
            e.Graphics.FillRectangle(Brushes.LightGray, 500, 715, 330, 15);
            e.Graphics.DrawLine(pen, 600, 725, 600, 830);
            e.Graphics.DrawLine(pen, 780, 725, 780, 830);

            //FIN

            e.Graphics.DrawLine(pen, 10, 830, 830, 830);

        }

        public void ImprimirDatos(Int64 numPedido,string strImpresora ,out string strMensaje,out bool logrespuesta)
        {
            try
            {

                var parametros = new List<SqlParameter> {new SqlParameter("numPedido", numPedido)};
                var dts = _dataHelper.EjecutarSp<DataSet>("fc_spImprimirFactura", parametros);


                if (dts != null)
                {
                    if (dts.Tables.Count > 0)
                    {
                        _dts = dts;

                        DataTable altDetalle = new DataTable();

                        altDetalle.Columns.Add("codigo", typeof(string));
                        altDetalle.Columns.Add("descripcion", typeof(string));
                        altDetalle.Columns.Add("cantidad", typeof(int));
                        altDetalle.Columns.Add("valorunitario", typeof(Int32));
                        altDetalle.Columns.Add("valortotal", typeof(Int32));


                        var dtdetalle = dts.Tables[1];

                        double totalpag = (double) dtdetalle.Rows.Count / 42;

                       _totalPaginas =(int) Math.Ceiling(totalpag);



                        int conteo = 0;

                        for (int i = 0; i < dtdetalle.Rows.Count; i++)
                        {
                            conteo++;
                            var row = altDetalle.NewRow();

                            row["codigo"] = dtdetalle.Rows[i]["codigo"].ToString();
                            row["descripcion"] = dtdetalle.Rows[i]["descripcion"].ToString();
                            row["cantidad"] = Convert.ToInt16(dtdetalle.Rows[i]["cantidad"]);
                            row["valorunitario"] = Convert.ToInt32(dtdetalle.Rows[i]["valorunitario"]);
                            row["valortotal"] = Convert.ToInt32(dtdetalle.Rows[i]["valortotal"]);


                            altDetalle.Rows.Add(row);


                            if (conteo == 42 || i == dtdetalle.Rows.Count - 1) 
                            {
                                _numPagina++;
                                _detalle = altDetalle;
                                conteo = 0;
                               
                                ////altDetalle.Clear();
                                ////_print.DefaultPageSettings.PaperSize = new PaperSize("letter", 850, 1100);
                                //_print.PrinterSettings.Duplex = Duplex.Vertical;
                                ////_print.PrinterSettings.PrinterName = strImpresora;
                                _print.Print();
                                ////_post.Print();
                                altDetalle.Clear();
                                paginaActual = 0;
                            }
                        }





                        strMensaje = "";
                        logrespuesta=true;
                    }
                    else
                    {
                        strMensaje = "No hay Datos Para Cargar";
                        logrespuesta= false;
                    }
                }
                else
                {
                    strMensaje = "No hay Datos Para Cargar";
                    logrespuesta= false;
                }

            }
            catch (Exception ex)
            {

                strMensaje = ex.Message;
                logrespuesta= false;
            }
        }

        void _print_PrintPage(object sender, PrintPageEventArgs e)
        {
            if (paginaActual == 0)
            {
                ImprimirPagina1(e); 
            }
            else 
            {
                ImprimirPagina2(e);
            }

            e.HasMorePages = paginaActual < 1;
            paginaActual++;

        }

        private void ImprimirPagina2(PrintPageEventArgs e)
        {

            var dtEncabezado = _dts.Tables[0];
            var dtInformativo = _dts.Tables[4];
            var dtPremios = _dts.Tables[3];

            int fontsize = 8;
            var xbrushes = Brushes.Black;
            var fueteTitulo = new Font("Tahoma", 9, FontStyle.Bold);
            var fuenteNormal = new Font("Tahoma", fontsize);
            var fuentepequeña = new Font("Tahoma", 5, FontStyle.Bold);
            var fuenteDetalle = new Font("Tahoma", 7, FontStyle.Bold);
            var fuentecolilla = new Font("Tahoma", 8, FontStyle.Bold);
            var fuentelistaDetalle = new Font("Arial", 7);
            var fuenteBancos = new Font("Arial", 5);



            //string rutaImagen = Properties.Settings.Default.Carpeta_Imagen.Trim();
            Image logo = Properties.Resources.Logo_Dolce;
            //Image logo = Image.FromFile(string.Format("{0}Logo Dolce.jpg", rutaImagen));

            e.Graphics.DrawImage(logo, 270, 5);
            e.Graphics.DrawString(dtEncabezado.Rows[0]["nombreempresa"].ToString().Trim(), fueteTitulo, xbrushes, 350, 10);
            e.Graphics.DrawString(string.Format("Nit:{0}", dtEncabezado.Rows[0]["nitempresa"]).ToString().Trim(), fueteTitulo, xbrushes, 350, 25);
            e.Graphics.DrawString(dtEncabezado.Rows[0]["ventadirecta"].ToString().Trim(), fueteTitulo, xbrushes, 350, 40);
            e.Graphics.DrawString(string.Format("FACTURA DE VENTA No {0}-{1}", dtEncabezado.Rows[0]["Prefijo"].ToString().Trim(), dtEncabezado.Rows[0]["numFactura"].ToString().Trim()), fueteTitulo, xbrushes, 500, 10);
            e.Graphics.DrawString(string.Format("PEDIDO No {0}", dtEncabezado.Rows[0]["numPedido"].ToString().Trim()), fueteTitulo, xbrushes, 500, 25);
            e.Graphics.DrawString(string.Format("ASESORA: {0}",dtEncabezado.Rows[0]["Asesora"].ToString()),fuentelistaDetalle,xbrushes,500,40);
            e.Graphics.DrawString("CAMPAÑA", fueteTitulo, xbrushes, 730, 10);
            e.Graphics.DrawString(dtEncabezado.Rows[0]["StrCampaña"].ToString().Trim(), fueteTitulo, xbrushes, 730, 25);

            //Caja del Numero de pedido zona y total de prendas

            e.Graphics.FillRectangle(Brushes.LightGray, 10, 10, 250, 40);

            e.Graphics.DrawString("PEDIDOS No", fuenteBancos, xbrushes, 15, 10);
            e.Graphics.DrawString("ZONA", fuenteBancos, xbrushes, 110, 10);
            e.Graphics.DrawString("TOTAL PRENDAS", fuenteBancos, xbrushes, 170, 10);

            var listaTotalInformacion = dtInformativo.Select("strtipo=7");

            e.Graphics.DrawString(listaTotalInformacion[0]["strObservacion"].ToString().Trim(), new Font("Arial", 20, FontStyle.Bold), xbrushes, 15, 20);
            e.Graphics.DrawString(listaTotalInformacion[0]["strCodigo"].ToString().Trim(), new Font("Arial", 20, FontStyle.Bold), xbrushes, 90, 20);
            e.Graphics.DrawString(listaTotalInformacion[0]["intCantidad"].ToString().Trim(), new Font("Arial", 20, FontStyle.Bold), xbrushes, 180, 20);

            //Fin caja 

            //Detalles de la Pagina posteriror

            var pen = new Pen(Color.LightGray) { DashStyle = DashStyle.Dot };


            e.Graphics.FillRectangle(Brushes.LightGray, 10, 55, 820, 15);
            e.Graphics.DrawString("INFORMATIVO DOLCE", fueteTitulo, xbrushes, 15, 55);


            //Lineas Divisorias
            //Premios Entragados

            e.Graphics.DrawString("PREMIOS ENTREGADOS", fuentecolilla, xbrushes, 15, 70);

            int conteo = 75;

            for (int i = 0; i < dtPremios.Rows.Count; i++)
            {
                conteo += 10;
                e.Graphics.DrawString(string.Format("{0} CANTIDAD {1}", dtPremios.Rows[i]["strCodigo"].ToString().Trim(), dtPremios.Rows[i]["intCantidad"]), fuentelistaDetalle, xbrushes, 15, conteo);
            }

            e.Graphics.DrawLine(pen, 500, 60, 500, 830);

            //Puntos Acumulados

            var listaPuntos = dtInformativo.Select("strTipo=2");

            e.Graphics.DrawString("PUNTOS ACUMULADOS", fuentecolilla, xbrushes, 15, 255);

            conteo = 260;

            for (int i = 0; i < listaPuntos.Count(); i++)
            {
                if (i == 0)
                {
                    conteo += 10;
                    e.Graphics.DrawString(listaPuntos[i]["strObservacion"].ToString().Trim(), fuenteBancos, xbrushes, 15, conteo);
                }
                else if (i == 1)
                {
                    conteo += 15;
                    e.Graphics.DrawString(listaPuntos[i]["strObservacion"].ToString().Trim(), fuenteBancos, xbrushes, 15, conteo);
                }
                else
                {
                    conteo += 10;
                    e.Graphics.DrawString(listaPuntos[i]["strObservacion"].ToString().Trim(), fuenteBancos, xbrushes, 15, conteo);    
                }
                
            }
            e.Graphics.DrawLine(pen, 10, 250, 500, 250);


            //Recordatorio Otros

            var listarecordatorio = dtInformativo.Select("strTipo=3");
            e.Graphics.DrawString("OTROS", fuentecolilla, xbrushes, 15, 405);

            conteo = 415;

            if (listarecordatorio.Count() > 0)
            {

                for (int i = 0; i < listarecordatorio.Count(); i++)
                {
                    string caracter = listarecordatorio[i]["strObservacion"].ToString().Trim().Substring(1,1);
                   
                    if (caracter == "N")
                    {
                        string textoObservacion = listarecordatorio[i]["strObservacion"].ToString().Trim().Substring(3);
                        conteo += 15;
                        e.Graphics.DrawString(textoObservacion, new Font("Arial",8,FontStyle.Bold), xbrushes, 15, conteo);
                    }
                    else
                    {
                        conteo += 10;
                        e.Graphics.DrawString(listarecordatorio[i]["strObservacion"].ToString().Trim(), fuenteBancos, xbrushes, 15, conteo);
                    }
                   
                }
            }


            //Imagen Publicidad Dolce

            Image publicidad = Properties.Resources.publicidad_dolce;
            e.Graphics.DrawImage(publicidad, 20, 520);

            //linea de Otros 

            e.Graphics.DrawLine(pen, 10, 380, 500, 380);

            e.Graphics.DrawLine(pen, 10, 500, 500, 500);

            //Productos Agotados
            e.Graphics.FillRectangle(Brushes.LightGray, 500, 85, 330, 15);
            e.Graphics.DrawString("PRODUCTOS AGOTADOS",fuentecolilla,xbrushes,500,85);
            
            e.Graphics.DrawLine(pen, 600, 90, 600, 525);
            e.Graphics.DrawLine(pen, 780, 90, 780, 525);
            e.Graphics.DrawLine(pen, 500, 525, 830, 525);

            var listaAgotados = dtInformativo.Select("strTipo=4");

            conteo = 95;
            
            if (listaAgotados.Count() > 0)
            {
                for (int i = 0; i < listaAgotados.Count(); i++)
                {
                    conteo += 10;
                    e.Graphics.DrawString(listaAgotados[i]["strCodigo"].ToString(), fuenteBancos, xbrushes, 505, conteo);
                    e.Graphics.DrawString(listaAgotados[i]["strObservacion"].ToString(), fuenteBancos, xbrushes, 605, conteo);
                    e.Graphics.DrawString(listaAgotados[i]["intCantidad"].ToString(), fuenteBancos, xbrushes, 800, conteo);
                }
            }



            //Cambios Surtidos


            e.Graphics.FillRectangle(Brushes.LightGray, 500, 540, 330, 15);
            e.Graphics.DrawString("CAMBIOS SURTIDOS",fuentecolilla,xbrushes,500,540);
            e.Graphics.DrawLine(pen, 600, 550, 600, 700);
            e.Graphics.DrawLine(pen, 780, 550, 780, 700);
            e.Graphics.DrawLine(pen, 500, 700, 830, 700);

            var listaCambiosSurtidos = dtInformativo.Select("strTipo=5");
            
            conteo = 560;

            if (listaCambiosSurtidos.Count() > 0)
            {
                for (int i = 0; i < listaCambiosSurtidos.Count(); i++)
                {
                    conteo=conteo + 10;
                    e.Graphics.DrawString(listaCambiosSurtidos[i]["strCodigo"].ToString(), fuenteBancos, xbrushes, 505, conteo);
                    e.Graphics.DrawString(listaCambiosSurtidos[i]["strObservacion"].ToString(), fuenteBancos, xbrushes, 605, conteo);
                    e.Graphics.DrawString(listaCambiosSurtidos[i]["intCantidad"].ToString(), fuenteBancos, xbrushes, 800, conteo);
                }
            }




            //Cambios Agotados
            e.Graphics.FillRectangle(Brushes.LightGray, 500, 715, 330, 15);
            e.Graphics.DrawString("CAMBIOS AGOTADOS",fuentecolilla,xbrushes,500,715);
            e.Graphics.DrawLine(pen, 600, 725, 600, 830);
            e.Graphics.DrawLine(pen, 780, 725, 780, 830);

            var listaCambiosAgotados = dtInformativo.Select("strTipo=6");

            conteo = 725;

            if (listaCambiosAgotados.Count() > 0)
            {
                for (int i = 0; i < listaCambiosAgotados.Count(); i++)
                {
                    conteo = conteo + 10;
                    e.Graphics.DrawString(listaCambiosAgotados[i]["strCodigo"].ToString(), fuenteBancos, xbrushes, 505, conteo);
                    e.Graphics.DrawString(listaCambiosAgotados[i]["strObservacion"].ToString(), fuenteBancos, xbrushes, 605, conteo);
                    e.Graphics.DrawString(listaCambiosAgotados[i]["intCantidad"].ToString(), fuenteBancos, xbrushes, 800, conteo);
                }
            }


            //FIN

            e.Graphics.DrawLine(pen, 10, 830, 830, 830);


            //Detalle de Puntos entregados

            var dtpuntos = _dts.Tables[5];

            if (dtpuntos.Rows.Count > 0)
            {

                int Acumulados = Convert.ToInt32(dtpuntos.Rows[0]["Puntos_Acumulados"]);
                int total = Convert.ToInt32(dtpuntos.Rows[0]["Saldo"]);
                int devoluciones = Convert.ToInt32(dtpuntos.Rows[0]["Devoluciones"]);
                int utilizados = Convert.ToInt32(dtpuntos.Rows[0]["Puntos_Utilizados"]);
                int pendientes = Convert.ToInt32(dtpuntos.Rows[0]["Puntos_Pendientes"]);


                int saldo = ((Acumulados - utilizados) - devoluciones);

                Image anitaImage = Properties.Resources.Anita;
                e.Graphics.DrawImage(anitaImage, 20, 850);
                e.Graphics.DrawString("EXTRACTO DE", new Font("Arial", 12, FontStyle.Bold), xbrushes, 20, 990);
                e.Graphics.DrawString("PUNTOS", new Font("Arial", 12, FontStyle.Bold), xbrushes, 20, 1005);

                e.Graphics.DrawRectangle(new Pen(Brushes.Black), 200, 850, 300, 200);
                //e.Graphics.DrawLine(new Pen(Brushes.Black), 400, 850,400, 1050);
                e.Graphics.FillRectangle(Brushes.LightGray, 201, 851, 298, 19);

                e.Graphics.DrawLine(new Pen(Brushes.Black), 200, 870, 500, 870);
                e.Graphics.DrawLine(new Pen(Brushes.Black), 200, 890, 500, 890);
                e.Graphics.DrawLine(new Pen(Brushes.Black), 200, 910, 500, 910);
                e.Graphics.DrawLine(new Pen(Brushes.Black), 200, 930, 500, 930);
                e.Graphics.DrawLine(new Pen(Brushes.Black), 200, 950, 500, 950);

                e.Graphics.DrawLine(new Pen(Brushes.Black), 400, 870, 400, 930);

                e.Graphics.DrawString("Acumulacion Puntos", new Font("Arial", 8, FontStyle.Bold), xbrushes, 205, 855);
                e.Graphics.DrawString("Puntos por Venta", new Font("Arial", 8, FontStyle.Bold), xbrushes, 205, 875);
                e.Graphics.DrawString("Puntos por Referido", new Font("Arial", 8, FontStyle.Bold), xbrushes, 205, 895);
                e.Graphics.DrawString("Puntos por Regalo", new Font("Arial", 8, FontStyle.Bold), xbrushes, 205, 915);

                //Datos de Puntos Acumulados 

                Int32 puntosAcumulados =Convert.ToInt32(dtpuntos.Rows[0]["Puntos_Acumulados"]);
                Int32 puntosReferidos = Convert.ToInt32(dtpuntos.Rows[0]["Puntos_Referidos"]);
                Int32 puntosRegalo =Convert.ToInt32(dtpuntos.Rows[0]["Puntos_Regalo"]);
                Int32 puntosSaldo = Convert.ToInt32(dtpuntos.Rows[0]["Saldo"]);
                Int32 xdevoluciones=Convert.ToInt32(dtpuntos.Rows[0]["Devoluciones"]);
                Int32 anulados =Convert.ToInt32(dtpuntos.Rows[0]["Puntos_Anulados"]);
                Int32 puntosUtilizados = Convert.ToInt32(dtpuntos.Rows[0]["Puntos_Utilizados"]);
                Int32 puntosPendientes = Convert.ToInt32(dtpuntos.Rows[0]["Puntos_Pendientes"]);

                Int32 puntosVenta = puntosAcumulados - (puntosReferidos + puntosRegalo);


                e.Graphics.DrawString(puntosVenta.ToString(), new Font("Arial",10, FontStyle.Bold), xbrushes,420, 873);
                e.Graphics.DrawString(puntosReferidos.ToString(), new Font("Arial", 10, FontStyle.Bold), xbrushes, 420, 893);
                e.Graphics.DrawString(puntosRegalo.ToString(), new Font("Arial", 10, FontStyle.Bold), xbrushes, 420, 913);

                
                e.Graphics.FillRectangle(Brushes.LightGray, 201, 931, 298, 18);
                e.Graphics.DrawLine(new Pen(Brushes.Black), 200, 970, 500, 970);
                e.Graphics.DrawLine(new Pen(Brushes.Black), 200, 990, 500, 990);
                e.Graphics.DrawLine(new Pen(Brushes.Black), 200, 1010, 500, 1010);

                e.Graphics.DrawLine(new Pen(Brushes.Black), 400, 950, 400, 1010);

                e.Graphics.DrawString("Deduccion de Puntos", new Font("Arial", 8, FontStyle.Bold), xbrushes, 205, 935);
                e.Graphics.DrawString("Puntos Anulados", new Font("Arial", 8, FontStyle.Bold), xbrushes, 205, 955);
                e.Graphics.DrawString("Puntos por Devolucion", new Font("Arial", 8, FontStyle.Bold), xbrushes, 205, 975);
                e.Graphics.DrawString("Puntos Utilizados", new Font("Arial", 8, FontStyle.Bold), xbrushes, 205, 995);


                e.Graphics.DrawString(string.Format("-{0}",anulados), new Font("Arial", 10, FontStyle.Bold), xbrushes, 420, 953);
                e.Graphics.DrawString(string.Format("-{0}",xdevoluciones), new Font("Arial", 10, FontStyle.Bold), xbrushes, 420, 973);
                e.Graphics.DrawString(string.Format("-{0}",puntosUtilizados), new Font("Arial", 10, FontStyle.Bold), xbrushes, 420, 993);


                e.Graphics.DrawString("SALDO DE PUNTOS", new Font("Arial", 10, FontStyle.Bold), xbrushes, 205, 1021);
                e.Graphics.DrawString(puntosSaldo.ToString(), new Font("Arial", 14, FontStyle.Bold), xbrushes, 405, 1017);

                //Puntos Pendientes
                e.Graphics.DrawString("PUNTOS ACUMULADOS", new Font("Arial", 11, FontStyle.Bold), xbrushes, 520, 850);
                e.Graphics.DrawString("EN ESTA CAMPAÑA", new Font("Arial", 11, FontStyle.Bold), xbrushes, 520, 870);

                e.Graphics.DrawString(puntosPendientes.ToString() , new Font("Arial", 18, FontStyle.Bold), xbrushes, 520,900);

                e.Graphics.DrawString("¡SOLO SI HACES TU PAGO A TIEMPO!", new Font("Arial", 11, FontStyle.Bold), xbrushes, 520, 950);

                //e.Graphics.DrawString(string.Format("PUNTOS ACUMULADOS A CAMPAÑA {0}", dtpuntos.Rows[0]["Campaña_anterior"]), new Font("Tahoma", 12, FontStyle.Bold), xbrushes, 200, 850);
                //e.Graphics.DrawString("PUNTOS", new Font("Tahoma", 10, FontStyle.Bold), xbrushes, 200, 870);
                //e.Graphics.DrawString("PUNTOS UTILIZADOS", new Font("Tahoma", 10, FontStyle.Bold), xbrushes, 300, 870);
                //e.Graphics.DrawString("DEVOLUCIONES", new Font("Tahoma", 10, FontStyle.Bold), xbrushes, 480, 870);
                //e.Graphics.DrawString("TOTAL", new Font("Tahoma", 10, FontStyle.Bold), xbrushes, 650, 870);
                //e.Graphics.DrawString("TOTAL", new Font("Tahoma", 10, FontStyle.Bold), xbrushes, 700, 870);

                //e.Graphics.DrawString(dtpuntos.Rows[0]["Puntos_Acumulados"].ToString(), fueteTitulo, xbrushes, 200, 890);
                //e.Graphics.DrawString(dtpuntos.Rows[0]["Puntos_Utilizados"].ToString(), fueteTitulo, xbrushes, 350, 890);
                //e.Graphics.DrawString(dtpuntos.Rows[0]["Devoluciones"].ToString(), fueteTitulo, xbrushes, 480, 890);
                //e.Graphics.DrawString(saldo.ToString(), fueteTitulo, xbrushes, 650, 890);
                //e.Graphics.DrawString(saldo.ToString(), fueteTitulo, xbrushes, 700, 890);


                //e.Graphics.DrawLine(pen, 200, 920, 920, 920);

                //e.Graphics.DrawString("PUNTOS ACUMULADOS EN ESTA CAMPAÑA", fueteTitulo, xbrushes, 200, 950);
                //e.Graphics.DrawString(pendientes.ToString(), new Font("Tahoma", 16, FontStyle.Bold), xbrushes, 600, 950);
                //e.Graphics.DrawString("SOLO SI HACES TU PAGO A TIEMPO", fueteTitulo, xbrushes, 200, 970);
                //e.Graphics.DrawString("¡EXITOS Y A REDIMIR MUCHOS PREMIOS!", new Font("Segoe Script", 14), xbrushes, 200, 1010);
            }
            else
            {
                var dtPendiente = _dts.Tables[6];


                Int32 puntosPendientes = Convert.ToInt32(dtPendiente.Rows[0]["Puntos_Pendientes"]);

                Image anitaImage = Properties.Resources.Anita;
                e.Graphics.DrawImage(anitaImage, 20, 850);
                e.Graphics.DrawString("EXTRACTO DE", new Font("Arial", 12, FontStyle.Bold), xbrushes, 20, 990);
                e.Graphics.DrawString("PUNTOS", new Font("Arial", 12, FontStyle.Bold), xbrushes, 20, 1005);

                e.Graphics.DrawRectangle(new Pen(Brushes.Black), 200, 850, 300, 200);
                //e.Graphics.DrawLine(new Pen(Brushes.Black), 400, 850,400, 1050);
                e.Graphics.FillRectangle(Brushes.LightGray, 201, 851, 298, 19);

                e.Graphics.DrawLine(new Pen(Brushes.Black), 200, 870, 500, 870);
                e.Graphics.DrawLine(new Pen(Brushes.Black), 200, 890, 500, 890);
                e.Graphics.DrawLine(new Pen(Brushes.Black), 200, 910, 500, 910);
                e.Graphics.DrawLine(new Pen(Brushes.Black), 200, 930, 500, 930);
                e.Graphics.DrawLine(new Pen(Brushes.Black), 200, 950, 500, 950);

                e.Graphics.DrawLine(new Pen(Brushes.Black), 400, 870, 400, 930);

                e.Graphics.DrawString("Acumulacion Puntos", new Font("Arial", 8, FontStyle.Bold), xbrushes, 205, 855);
                e.Graphics.DrawString("Puntos por Venta", new Font("Arial", 8, FontStyle.Bold), xbrushes, 205, 875);
                e.Graphics.DrawString("Puntos por Referido", new Font("Arial", 8, FontStyle.Bold), xbrushes, 205, 895);
                e.Graphics.DrawString("Puntos por Regalo", new Font("Arial", 8, FontStyle.Bold), xbrushes, 205, 915);

                //Datos de Puntos Acumulados 


                e.Graphics.DrawString("0", new Font("Arial", 10, FontStyle.Bold), xbrushes, 420, 873);
                e.Graphics.DrawString("0", new Font("Arial", 10, FontStyle.Bold), xbrushes, 420, 893);
                e.Graphics.DrawString("0", new Font("Arial", 10, FontStyle.Bold), xbrushes, 420, 913);


                e.Graphics.FillRectangle(Brushes.LightGray, 201, 931, 298, 18);
                e.Graphics.DrawLine(new Pen(Brushes.Black), 200, 970, 500, 970);
                e.Graphics.DrawLine(new Pen(Brushes.Black), 200, 990, 500, 990);
                e.Graphics.DrawLine(new Pen(Brushes.Black), 200, 1010, 500, 1010);

                e.Graphics.DrawLine(new Pen(Brushes.Black), 400, 950, 400, 1010);

                e.Graphics.DrawString("Deduccion de Puntos", new Font("Arial", 8, FontStyle.Bold), xbrushes, 205, 935);
                e.Graphics.DrawString("Puntos Anulados", new Font("Arial", 8, FontStyle.Bold), xbrushes, 205, 955);
                e.Graphics.DrawString("Puntos por Devolucion", new Font("Arial", 8, FontStyle.Bold), xbrushes, 205, 975);
                e.Graphics.DrawString("Puntos Utilizados", new Font("Arial", 8, FontStyle.Bold), xbrushes, 205, 995);


                e.Graphics.DrawString(string.Format("-{0}", 0), new Font("Arial", 10, FontStyle.Bold), xbrushes, 420, 953);
                e.Graphics.DrawString(string.Format("-{0}", 0), new Font("Arial", 10, FontStyle.Bold), xbrushes, 420, 973);
                e.Graphics.DrawString(string.Format("-{0}", 0), new Font("Arial", 10, FontStyle.Bold), xbrushes, 420, 993);


                e.Graphics.DrawString("SALDO DE PUNTOS", new Font("Arial", 10, FontStyle.Bold), xbrushes, 205, 1021);
                e.Graphics.DrawString("0", new Font("Arial", 14, FontStyle.Bold), xbrushes, 405, 1017);

                //Puntos Pendientes
                e.Graphics.DrawString("PUNTOS ACUMULADOS", new Font("Arial", 11, FontStyle.Bold), xbrushes, 520, 850);
                e.Graphics.DrawString("EN ESTA CAMPAÑA", new Font("Arial", 11, FontStyle.Bold), xbrushes, 520, 870);

                e.Graphics.DrawString(puntosPendientes.ToString(), new Font("Arial", 18, FontStyle.Bold), xbrushes, 520, 900);

                e.Graphics.DrawString("¡SOLO SI HACES TU PAGO A TIEMPO!", new Font("Arial", 11, FontStyle.Bold), xbrushes, 520, 950);


            }
            
          

        }


        private void ImprimirPagina1(PrintPageEventArgs e)
        {
            var dtEncabezado = _dts.Tables[0];
            //var dtDetalle = _dts.Tables[1];
            var dtDetalle = _detalle;
            var dtBancos = _dts.Tables[2];
            var dtPremios = _dts.Tables[3];
            //var dtSaldosAsesora = _dts.Tables[7];
            //var dtSaldoAFavor = _dts.Tables[7];
            //var dtSaldoPendiente = _dts.Tables[8];

             Int32 numSAldoAfavor=0;
             Int32 numSaldoPendiente = 0;
             Int32 totalpagar = 0;

             //if (dtSaldosAsesora.Rows.Count > 0) 
             //{
             //    numSAldoAfavor = Convert.ToInt32(dtSaldosAsesora.Rows[0]["curSaldoAfavor"]);
             //    numSaldoPendiente = Convert.ToInt32(dtSaldosAsesora.Rows[0]["curSaldoEnContra"]);
             //}

           


            //strTipo=4 Productos Agotados
            //strTipo=5 Cambios Surtidos
            //strTipo=6 Cambios Agotados
            //strTipo=
            var dtOtros = _dts.Tables[4];

            //string rutaImagen = Properties.Settings.Default.Carpeta_Imagen.Trim();


            Image logo = Properties.Resources.Logo_Dolce_1;
            Image logo1 = Properties.Resources.Logo_Dolce;
            //Image logo = Image.FromFile(string.Format("{0}Logo_Dolce.jpg", rutaImagen));


            int fontsize = 8;
            var xbrushes = Brushes.Black;
            var fueteTitulo = new Font("Tahoma", 9, FontStyle.Bold);
            var fuenteNormal = new Font("Tahoma", fontsize);
            var fuentepequeña = new Font("Tahoma", 5, FontStyle.Bold);
            var fuenteDetalle = new Font("Tahoma", 7, FontStyle.Bold);
            var fuentecolilla = new Font("Tahoma", 8, FontStyle.Bold);
            var fuentelistaDetalle = new Font("Arial", 8);
            var fuenteBancos = new Font("Arial", 5);

            e.Graphics.DrawImage(logo, 10, 10);

            //Paginacion 
            e.Graphics.DrawString(string.Format("Pagina {0} de {1}", _numPagina, _totalPaginas), new Font("Arial", 7, FontStyle.Bold), xbrushes, 700, 10);
            //Fin Paginacion

            //Datos del Encabezado

            //e.Graphics.DrawString("D O L C E",new Font("Arial",10,FontStyle.Bold),xbrushes,80,20);
            //e.Graphics.DrawString("Un mundo de moda.", new Font("Arial", 10,FontStyle.Bold), xbrushes, 80, 40);

            e.Graphics.DrawString(dtEncabezado.Rows[0]["nombreempresa"].ToString(), fueteTitulo, xbrushes, 250, 25);
            e.Graphics.DrawString(string.Format("Nit:{0}", dtEncabezado.Rows[0]["nitempresa"]), fueteTitulo, xbrushes, 250, 40);
            e.Graphics.DrawString(dtEncabezado.Rows[0]["ventadirecta"].ToString(), fueteTitulo, xbrushes, 250, 55);
            e.Graphics.DrawString(dtEncabezado.Rows[0]["regimen"].ToString(), fuentepequeña, xbrushes, 250, 75);
            //e.Graphics.DrawString(dtEncabezado.Rows[0]["contribuyentes"].ToString(), fuentepequeña, xbrushes, 250, 75);
            //e.Graphics.DrawString(dtEncabezado.Rows[0]["resolucionGC"].ToString(), fuentepequeña, xbrushes, 250, 90);

            e.Graphics.DrawString("FACTURA DE VENTA No", fueteTitulo, xbrushes, 470, 20);
            e.Graphics.DrawString(
                string.Format("{0}-{1}", dtEncabezado.Rows[0]["Prefijo"].ToString().Trim(),
                    dtEncabezado.Rows[0]["numFactura"].ToString().Trim()), fueteTitulo, xbrushes, 650, 20);
            e.Graphics.DrawString("PEDIDO No", fueteTitulo, xbrushes, 470, 35);
            e.Graphics.DrawString(dtEncabezado.Rows[0]["numPedido"].ToString(), fueteTitulo, xbrushes, 650, 35);
            e.Graphics.DrawString("FECHA FACTURA", fueteTitulo, xbrushes, 470, 50);
            e.Graphics.DrawString(Convert.ToDateTime(dtEncabezado.Rows[0]["fechafactura"]).Date.ToShortDateString(), fueteTitulo,
                xbrushes, 650, 50);
            e.Graphics.DrawString("FECHA VENCIMIENTO", fueteTitulo, xbrushes, 470, 65);
            e.Graphics.DrawString(Convert.ToDateTime(dtEncabezado.Rows[0]["fechavence"]).Date.ToShortDateString(), fueteTitulo,
                xbrushes, 650, 65);

            e.Graphics.DrawString("CAMPAÑA", fueteTitulo, xbrushes, 750, 50);
            e.Graphics.DrawString(dtEncabezado.Rows[0]["strCampaña"].ToString().Trim(), fueteTitulo, xbrushes, 750, 65);

            //Datos del Cliente


            e.Graphics.FillRectangle(Brushes.LightGray, 10, 90, 830, 70);

            e.Graphics.DrawString("Cliente", fueteTitulo, xbrushes, 20, 95);
            e.Graphics.DrawString(dtEncabezado.Rows[0]["Asesora"].ToString().Trim(), fuenteNormal, xbrushes, 80, 95);
            e.Graphics.DrawString("Cedula", fueteTitulo, xbrushes, 20, 110);
            e.Graphics.DrawString(dtEncabezado.Rows[0]["cedula"].ToString().Trim(), new Font("Tahoma",10,FontStyle.Bold), xbrushes, 80, 110);
            e.Graphics.DrawString("Zona", fueteTitulo, xbrushes, 20, 125);
            e.Graphics.DrawString(dtEncabezado.Rows[0]["Zona"].ToString().Trim(), fuenteNormal, xbrushes, 80, 125);
            e.Graphics.DrawString("Seccion", fueteTitulo, xbrushes, 20, 140);
            e.Graphics.DrawString(dtEncabezado.Rows[0]["Seccion"].ToString().Trim(), fuenteNormal, xbrushes, 80, 140);
            e.Graphics.DrawString("Telefono", fueteTitulo, xbrushes, 270, 110);
            e.Graphics.DrawString(dtEncabezado.Rows[0]["telefono"].ToString().Trim(), fuenteNormal, xbrushes, 370, 110);
            e.Graphics.DrawString("Direccion", fueteTitulo, xbrushes, 270, 125);
            e.Graphics.DrawString(dtEncabezado.Rows[0]["direccion"].ToString().Trim(), fuenteNormal, xbrushes, 370, 125);
            e.Graphics.DrawString("Ciudad/Barrio", fueteTitulo, xbrushes, 270, 140);
            e.Graphics.DrawString(
                string.Format("{0}/{1}", dtEncabezado.Rows[0]["ciudad"].ToString().Trim(),
                    dtEncabezado.Rows[0]["barrio"].ToString().Trim()), fuenteNormal, xbrushes, 370, 140);
            e.Graphics.DrawString("Directora", fueteTitulo, xbrushes, 580, 95);
            e.Graphics.DrawString(dtEncabezado.Rows[0]["directora"].ToString().Trim(), fuenteNormal, xbrushes, 650, 95);
            e.Graphics.DrawString("Celular", fueteTitulo, xbrushes, 580, 110);
            e.Graphics.DrawString(dtEncabezado.Rows[0]["celulardirectora"].ToString().Trim(), fuenteNormal, xbrushes, 650, 110);
            e.Graphics.DrawString("Telefono", fueteTitulo, xbrushes, 580, 125);
            e.Graphics.DrawString(dtEncabezado.Rows[0]["telefonodirectora"].ToString().Trim(), fuenteNormal, xbrushes, 650, 125);

            //Fin Datos del Cliente

            //Encabezado Columnas

            e.Graphics.FillRectangle(Brushes.LightGray, 10, 170, 830, 15);

            e.Graphics.DrawString("CODIGO", fuenteDetalle, xbrushes, 20, 172);
            e.Graphics.DrawString("DESCRIPCION", fuenteDetalle, xbrushes, 210, 172);
            e.Graphics.DrawString("CANTIDAD", fuenteDetalle, xbrushes, 430, 172);
            e.Graphics.DrawString("VALOR UNITARIO", fuenteDetalle, xbrushes, 560, 172);
            e.Graphics.DrawString("VALOR TOTAL", fuenteDetalle, xbrushes, 730, 172);

            //Fin Encabezado Columnas

            //Lineas verticales del Detalle

            var pen = new Pen(Color.LightGray, 1) { DashStyle = DashStyle.Dot };

            e.Graphics.DrawLine(pen, 90, 180, 90, 650);
            e.Graphics.DrawLine(pen, 400, 180, 400, 650);
            e.Graphics.DrawLine(pen, 520, 180, 520, 650);
            e.Graphics.DrawLine(pen, 700, 180, 700, 650);
            //Fin Lineas verticales del Detalle

            //Linea Final Del Detalle

            e.Graphics.DrawLine(pen, 10, 650, 830, 650);

            //Fin de Linea Final Detalle

            //Detalle de Productos

            int conteo = 185;
            int linea = 186;

            //for (int i = 0; i < 42; i++)
            //{
            //    e.Graphics.DrawString("0000", fuentelistaDetalle, xbrushes, 20, conteo);
            //    e.Graphics.DrawString("SHOR CONTROL SBELTA REF 5946", fuentelistaDetalle, xbrushes, 95,
            //        conteo);
            //    e.Graphics.DrawString("1", fuentelistaDetalle, xbrushes, 450, conteo);
            //    e.Graphics.DrawString("$ 1.500",
            //        fuentelistaDetalle, xbrushes, 550, conteo);
            //    e.Graphics.DrawString("$ 1.500", fuentelistaDetalle,
            //        xbrushes, 720, conteo);

            //    conteo = conteo + 11;

            //    linea = linea + 11;

            //    e.Graphics.DrawLine(pen, 10, linea, 830, linea);
            //}

            for (int i = 0; i < dtDetalle.Rows.Count; i++)
            {
                e.Graphics.DrawString(dtDetalle.Rows[i]["codigo"].ToString().Trim(), fuentelistaDetalle, xbrushes, 20, conteo);
                e.Graphics.DrawString(dtDetalle.Rows[i]["descripcion"].ToString().Trim(), fuentelistaDetalle, xbrushes, 95,
                    conteo);
                e.Graphics.DrawString(dtDetalle.Rows[i]["cantidad"].ToString().Trim(), fuentelistaDetalle, xbrushes, 450, conteo);
                e.Graphics.DrawString(string.Format("$ {0:##,##}", Convert.ToInt64(dtDetalle.Rows[i]["valorunitario"])),
                    fuentelistaDetalle, xbrushes, 550, conteo);
                e.Graphics.DrawString(string.Format("$ {0:##,##}", dtDetalle.Rows[i]["valortotal"]), fuentelistaDetalle,
                    xbrushes, 720, conteo);

                conteo = conteo + 11;

                linea = linea + 11;

                e.Graphics.DrawLine(pen, 10, linea, 830, linea);
            }


            //Fin Detalle de Productos

            //Observacion

            e.Graphics.DrawString("OBSERVACION", fueteTitulo, xbrushes, 20, 745);
            e.Graphics.DrawString(dtEncabezado.Rows[0]["Observacion2"].ToString().Trim(), fuenteNormal, xbrushes, 20, 765);


            //Fin Observacion

            //Total Factura

            if (_numPagina == _totalPaginas) 
            {
                
                Int32 totalFactura = Convert.ToInt32(dtEncabezado.Rows[0]["totalapagar"]);

                e.Graphics.FillRectangle(Brushes.LightGray, 550, 655, 280, 125);
                e.Graphics.DrawString("TOTAL VALOR FACTURA", fueteTitulo, xbrushes, 555, 655);
                e.Graphics.DrawString(string.Format("$ {0:##,##}", Convert.ToInt64(dtEncabezado.Rows[0]["totalapagar"])),
                    fueteTitulo, xbrushes, 730, 655);
                e.Graphics.DrawString("ESTA FACTURA INCLUYE IVA POR", new Font("Arial", 7), xbrushes, 555, 675);
                e.Graphics.DrawString(string.Format("$ {0:##,##}", Convert.ToInt64(dtEncabezado.Rows[0]["totaliva"])),
                    new Font("Arial", 8, FontStyle.Bold), xbrushes, 730, 675);
                //Fin Total Factura

                e.Graphics.DrawLine(new Pen(Brushes.Black) { DashStyle = DashStyle.Dot }, 553, 690, 825, 690);

                //Saldos a favor o encontra de la Asesora

                e.Graphics.DrawString("Saldo Pendiente", new Font("Arial", 10, FontStyle.Bold), xbrushes, 555, 700);
                e.Graphics.DrawString(string.Format("$ {0:##,##}", numSaldoPendiente), new Font("Arial", 8), xbrushes, 730, 700);

                e.Graphics.DrawString("Saldo a Favor", new Font("Arial", 10, FontStyle.Bold), xbrushes, 555, 725);
                e.Graphics.DrawString(string.Format("$ {0:##,##}", numSAldoAfavor), new Font("Arial", 8), xbrushes, 730, 725);

                e.Graphics.DrawLine(new Pen(Brushes.Black) { DashStyle = DashStyle.Dot }, 553, 745, 825, 745);

                totalpagar = (totalFactura + numSaldoPendiente) - numSAldoAfavor;

                e.Graphics.DrawString("TOTAL A PAGAR", fueteTitulo, xbrushes, 555, 750);
                e.Graphics.DrawString(string.Format("$ {0:##,##}", totalpagar), fueteTitulo, xbrushes, 730, 750);


                //Fin Saldos 
            }

            //Cuadro de Dialogo Pie de Pagina
            e.Graphics.DrawRectangle(Pens.Black, 10, 785, 819, 25);
            e.Graphics.DrawString(dtEncabezado.Rows[0]["Observacion"].ToString().Trim(), new Font("Arial", 6), xbrushes, 20, 787);
            e.Graphics.DrawString(dtEncabezado.Rows[0]["resolucion"].ToString().Trim(), new Font("Arial", 6), xbrushes, 20, 797);

            //Fin Cuadro de Dialogo

            //Colilla de Pago Izquierda
            e.Graphics.DrawImage(logo1, 10, 820);
            e.Graphics.DrawString(dtEncabezado.Rows[0]["nombreempresa"].ToString().Trim(), fuentecolilla, xbrushes, 100, 820);
            e.Graphics.DrawString(string.Format("Nit:{0}", dtEncabezado.Rows[0]["nitempresa"]).ToString().Trim(), fuentecolilla,
                xbrushes, 100, 835);
            e.Graphics.DrawString(dtEncabezado.Rows[0]["ventadirecta"].ToString().Trim(), fuentecolilla, xbrushes, 100, 850);
            e.Graphics.DrawString("DOLCE", new Font("Tahoma", 12, FontStyle.Bold), xbrushes, 10, 870);
            e.Graphics.DrawString("¡UN MUNDO DE MODA!", new Font("Tahoma",9, FontStyle.Bold), xbrushes, 10, 885);
            
            
            e.Graphics.DrawString("CUENTAS BANCARIAS", fuentecolilla, xbrushes, 10, 900);

            int conteoBanco = 900;
            for (int i = 0; i < dtBancos.Rows.Count; i++)
            {
                conteoBanco = conteoBanco + 12;
                e.Graphics.DrawString(dtBancos.Rows[i]["nombre"].ToString().Trim(), fuenteBancos, xbrushes, 10, conteoBanco);
                conteoBanco = conteoBanco + 12;
                e.Graphics.DrawString(dtBancos.Rows[i]["descripcion"].ToString().Trim(), fuenteBancos, xbrushes, 10, conteoBanco);
            }

            e.Graphics.DrawString("-COPIA BANCO-", fuentecolilla, xbrushes, 300, 820);

            e.Graphics.FillRectangle(Brushes.LightGray, 230, 850, 170, 40);

            e.Graphics.DrawString(dtEncabezado.Rows[0]["Asesora"].ToString().Trim(), fuentepequeña, xbrushes, 230, 850);
            e.Graphics.DrawString(string.Format("CEDULA {0}", dtEncabezado.Rows[0]["cedula"].ToString().Trim()),
                new Font("Arial", 5), xbrushes, 230, 860);
            e.Graphics.DrawString(string.Format("ZONA {0}", dtEncabezado.Rows[0]["Zona"].ToString().Trim()),
                new Font("Arial", 5), xbrushes, 230, 870);
            e.Graphics.DrawString("VALOR:", fueteTitulo, xbrushes, 270, 875);
            e.Graphics.DrawString(string.Format("$ {0:##,##}", totalpagar),
                fueteTitulo, xbrushes, 320, 875);
            e.Graphics.DrawString(
                string.Format("FECHA VENCIMIENTO {0}",
                    Convert.ToDateTime(dtEncabezado.Rows[0]["fechavence"]).Date.ToShortDateString()), new Font("Arial", 5),
                xbrushes, 270, 870);

            //Fin Colilla de Pago

            //Codigo de Barras

            e.Graphics.DrawString(dtEncabezado.Rows[0]["referenciabanco"].ToString().Trim(), fuentepequeña, xbrushes, 10, 960);
            //e.Graphics.DrawString(string.Format("*{0}*", dtEncabezado.Rows[0]["codbarras"].ToString().Trim()), new Font("Code128", 8), xbrushes, 10, 1010);


            var code128 = new Linear
            {
                Type = BarcodeType.CODE128,
                Data = dtEncabezado.Rows[0]["codbarras"].ToString().Trim(),
                LeftMargin = 0,
                RightMargin = 0
            };
            code128.drawBarcode("D:/Prueba.jpg");

            Image codbarras = Image.FromFile("D:/Prueba.jpg");


            e.Graphics.DrawImage(codbarras, 5, 970);
            e.Graphics.DrawImage(codbarras, 405, 970);

            //Fin Codigo de Barras

            //Colilla de Pago Derecha

            e.Graphics.DrawImage(logo1, 410, 820);
            e.Graphics.DrawString(dtEncabezado.Rows[0]["nombreempresa"].ToString().Trim(), fuentecolilla, xbrushes, 510, 820);
            e.Graphics.DrawString(string.Format("Nit:{0}", dtEncabezado.Rows[0]["nitempresa"]).ToString().Trim(), fuentecolilla,
                xbrushes, 510, 835);
            e.Graphics.DrawString(dtEncabezado.Rows[0]["ventadirecta"].ToString().Trim(), fuentecolilla, xbrushes, 510, 850);

            e.Graphics.DrawString("DOLCE", new Font("Tahoma", 12, FontStyle.Bold), xbrushes, 410, 870);
            e.Graphics.DrawString("¡UN MUNDO DE MODA!", new Font("Tahoma", 9, FontStyle.Bold), xbrushes, 410, 885);

            e.Graphics.DrawString("CUENTAS BANCARIAS", fuentecolilla, xbrushes, 410, 900);

            conteoBanco = 900;
            for (int i = 0; i < dtBancos.Rows.Count; i++)
            {
                conteoBanco = conteoBanco + 12;
                e.Graphics.DrawString(dtBancos.Rows[i]["nombre"].ToString().Trim(), fuenteBancos, xbrushes, 410, conteoBanco);
                conteoBanco = conteoBanco + 12;
                e.Graphics.DrawString(dtBancos.Rows[i]["descripcion"].ToString().Trim(), fuenteBancos, xbrushes, 410,
                    conteoBanco);
            }

            e.Graphics.DrawString("-COPIA DOLCE-", fuentecolilla, xbrushes, 710, 820);

            e.Graphics.FillRectangle(Brushes.LightGray, 650, 850, 170, 40);

            e.Graphics.DrawString(dtEncabezado.Rows[0]["Asesora"].ToString().Trim(), fuentepequeña, xbrushes, 650, 850);
            e.Graphics.DrawString(string.Format("CEDULA {0}", dtEncabezado.Rows[0]["cedula"].ToString().Trim()),
                new Font("Arial", 5), xbrushes, 650, 860);
            e.Graphics.DrawString(string.Format("ZONA {0}", dtEncabezado.Rows[0]["Zona"].ToString().Trim()),
                new Font("Arial", 5), xbrushes, 650, 870);
            e.Graphics.DrawString("VALOR:", fueteTitulo, xbrushes, 670, 875);
            e.Graphics.DrawString(string.Format("$ {0:##,##}", totalpagar),
                fueteTitulo, xbrushes, 720, 875);
            e.Graphics.DrawString(
                string.Format("FECHA VENCIMIENTO {0}",
                    Convert.ToDateTime(dtEncabezado.Rows[0]["fechavence"]).Date.ToShortDateString()), new Font("Arial", 5),
                xbrushes, 690, 870);

            e.Graphics.DrawString(dtEncabezado.Rows[0]["referenciabanco"].ToString().Trim(), fuentepequeña, xbrushes, 410, 960);

            //Fin Colilla de Pago
         
        }
    }
}
