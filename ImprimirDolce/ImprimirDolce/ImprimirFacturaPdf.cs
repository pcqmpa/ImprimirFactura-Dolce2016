using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Drawing.BarCodes;
using PdfSharp.Pdf;

namespace ImprimirDolce
{
    public class ImprimirFacturaPdf
    {
        readonly PdfDocument _pdf = new PdfDocument();
        DataSet _std= new DataSet();
        private readonly DataHelper _dataHelper;
        public ImprimirFacturaPdf()
        {
            _dataHelper= new DataHelper(new SqlConnectionStringBuilder(Properties.Settings.Default.conexion_dbDolce));
        }

        public void FacturaPdf(Int64 numeroPedido,string path, out string strMensaje, out bool logRespuesta,out string strArchivo)
        {
            var parametros = new List<SqlParameter> {new SqlParameter("numPedido", numeroPedido)};
            var dts = _dataHelper.EjecutarSp<DataSet>("fc_spImprimirFactura", parametros);

            if (dts != null)
            {
                if (dts.Tables.Count > 0)
                {
                    DataTable dtEncabezado = dts.Tables[0];
                    DataTable dtDetalle = dts.Tables[1];
                    DataTable dtInformativo = dts.Tables[4];
                    DataTable dtPremios = dts.Tables[3];
                    DataTable dtBancos = dts.Tables[2];
                    
                    XImage logo = Properties.Resources.Logo_Dolce;

                    _pdf.Info.Title = "Factura de venta Dolce S.A.S";

                    //Adicionamos la pagina pdf

                    PdfPage pgFrontal = _pdf.AddPage();
                    pgFrontal.Size = PageSize.Letter;


                    int fontsize = 8;

                    var fontNormal = new XFont("Tahoma", fontsize);
                    var fontTitulo = new XFont("Tahoma", 9, XFontStyle.Bold);

                    var pBrushes = XBrushes.Black;

                    XGraphics gfx = XGraphics.FromPdfPage(pgFrontal);

                    //Datos del encabezado

                    gfx.DrawImage(logo, 15, 20);
                    gfx.DrawString(dtEncabezado.Rows[0]["nombreempresa"].ToString().Trim(), fontTitulo, XBrushes.Black, 70, 20);
                    gfx.DrawString(string.Format("Nit:{0}", dtEncabezado.Rows[0]["nitempresa"]).ToString().Trim(), fontTitulo, XBrushes.Black, 70, 30);
                    gfx.DrawString(dtEncabezado.Rows[0]["ventadirecta"].ToString().Trim(), fontTitulo, XBrushes.Black, 70, 40);
                    gfx.DrawString(dtEncabezado.Rows[0]["regimen"].ToString().Trim(), new XFont("Tahoma", 5), XBrushes.Black, 70, 50);
                   // gfx.DrawString(dtEncabezado.Rows[0]["contribuyentes"].ToString().Trim(), new XFont("Tahoma", 5), XBrushes.Black, 70, 55);
                   // gfx.DrawString(dtEncabezado.Rows[0]["resolucionGC"].ToString().Trim(), new XFont("Tahoma", 5), XBrushes.Black, 70, 60);

                    //Datos de la Factura--------------------------------------------------------------------------------------------------------------------------------------------------------------

                    gfx.DrawString("FACTURA DE VENTA No", fontTitulo, XBrushes.Black, 390, 20);
                    gfx.DrawString(string.Format("{0}-{1}", dtEncabezado.Rows[0]["Prefijo"].ToString().Trim(), dtEncabezado.Rows[0]["numFactura"].ToString().Trim()), fontNormal, XBrushes.Black, 500, 20);
                    gfx.DrawString("PEDIDO No.", fontTitulo, XBrushes.Black, 390, 30);
                    gfx.DrawString(dtEncabezado.Rows[0]["numPedido"].ToString().Trim(), fontNormal, pBrushes, 500, 30);
                    gfx.DrawString("FECHA FACTURA", fontTitulo, pBrushes, 390, 40);
                    gfx.DrawString(dtEncabezado.Rows[0]["fechafactura"].ToString(), fontNormal, pBrushes, 500, 40);
                    gfx.DrawString("FECHA VENCIMINETO", fontTitulo, pBrushes, 390, 50);
                    gfx.DrawString(Convert.ToDateTime(dtEncabezado.Rows[0]["fechavence"]).Date.ToShortDateString(), fontNormal, pBrushes, 500, 50);

                    gfx.DrawString("CAMPAÑA", fontTitulo, pBrushes, 550, 40);
                    gfx.DrawString(dtEncabezado.Rows[0]["strCampaña"].ToString().Trim(), fontTitulo, pBrushes, 550, 50);

                    //Datos del CLiente

                    gfx.DrawRectangle(XBrushes.LightGray, 15, 70, 585, 45);

                    gfx.DrawString("Cliente", fontTitulo, pBrushes, 20, 80);
                    gfx.DrawString(dtEncabezado.Rows[0]["Asesora"].ToString().Trim(), fontNormal, pBrushes, 60, 80);
                    gfx.DrawString("Cedula", fontTitulo, pBrushes, 20, 90);
                    gfx.DrawString(dtEncabezado.Rows[0]["cedula"].ToString().Trim(), fontNormal, pBrushes, 60, 90);
                    gfx.DrawString("Zona", fontTitulo, pBrushes, 20, 100);
                    gfx.DrawString(dtEncabezado.Rows[0]["Zona"].ToString().Trim(), fontNormal, pBrushes, 60, 100);
                    gfx.DrawString("Seccion", fontTitulo, pBrushes, 20, 110);
                    gfx.DrawString(dtEncabezado.Rows[0]["Seccion"].ToString().Trim(), fontNormal, pBrushes, 60, 110);
                    gfx.DrawString("Telefono", fontTitulo, pBrushes, 200, 90);
                    gfx.DrawString(dtEncabezado.Rows[0]["telefono"].ToString().Trim(), fontNormal, pBrushes, 270, 90);
                    gfx.DrawString("Direccion", fontTitulo, pBrushes, 200, 100);
                    gfx.DrawString(dtEncabezado.Rows[0]["direccion"].ToString().Trim(), fontNormal, pBrushes, 270, 100);
                    gfx.DrawString("Ciudad/Barrio", fontTitulo, pBrushes, 200, 110);
                    gfx.DrawString(string.Format("{0}/{1}", dtEncabezado.Rows[0]["ciudad"].ToString().Trim(), dtEncabezado.Rows[0]["barrio"].ToString().Trim()), fontNormal, pBrushes, 270, 110);
                    gfx.DrawString("Directora", fontTitulo, pBrushes, 390, 80);
                    gfx.DrawString(dtEncabezado.Rows[0]["directora"].ToString().Trim(), fontNormal, pBrushes, 440, 80);
                    gfx.DrawString("Celular", fontTitulo, pBrushes, 390, 90);
                    gfx.DrawString(dtEncabezado.Rows[0]["celulardirectora"].ToString().Trim(), fontNormal, pBrushes, 440, 90);
                    gfx.DrawString("Telefono", fontTitulo, pBrushes, 390, 100);
                    gfx.DrawString(dtEncabezado.Rows[0]["telefonodirectora"].ToString().Trim(), fontNormal, pBrushes, 440, 100);

                    gfx.DrawRectangle(XBrushes.LightGray, 15, 120, 585, 10);

                    //Encabezado columnas

                    gfx.DrawString("CODIGO", new XFont("Arial", 7, XFontStyle.Bold), pBrushes, 20, 128);
                    gfx.DrawString("DESCRICPION", new XFont("Arial", 7, XFontStyle.Bold), pBrushes, 170, 128);
                    gfx.DrawString("CANTIDAD", new XFont("Arial", 7, XFontStyle.Bold), pBrushes, 320, 128);
                    gfx.DrawString("VALOR UNITARIO", new XFont("Arial", 7, XFontStyle.Bold), pBrushes, 420, 128);
                    gfx.DrawString("VALOR TOTAL", new XFont("Arial", 7, XFontStyle.Bold), pBrushes, 510, 128);

                    //Fin Datos del Encabezado---------------------------------------------------------------------------------------------------------------------------------------------------------------


                    gfx.DrawString("COPIA DE FACTURA", new XFont("Arial", 40, XFontStyle.BoldItalic), XBrushes.LightGray, 100, 300);

                    var fonDetalle = new XFont("Arial", 7);

                    //Detalle Factura------------------------------------------------------------------------------------------------------------------------------------------------------------------------

                    var pen = new XPen(XColors.LightGray, 0.7);
                    pen.DashStyle = XDashStyle.Dot;

                    int conteo = 130;
                    int linea = 132;

                    for (int i = 0; i < dtDetalle.Rows.Count; i++)
                    {
                        conteo = conteo + 10;
                        gfx.DrawString(dtDetalle.Rows[i]["codigo"].ToString().Trim(), fonDetalle, pBrushes, 20, conteo);
                        gfx.DrawString(dtDetalle.Rows[i]["descripcion"].ToString().Trim(), fonDetalle, pBrushes, 90, conteo);
                        gfx.DrawString(dtDetalle.Rows[i]["cantidad"].ToString().Trim(), fonDetalle, pBrushes, 330, conteo);
                        gfx.DrawString(string.Format("$ {0:##,##}", Convert.ToInt64(dtDetalle.Rows[i]["valorunitario"])), fonDetalle, pBrushes, 430, conteo);
                        gfx.DrawString(string.Format("$ {0:##,##}", dtDetalle.Rows[i]["valortotal"]), fonDetalle, pBrushes, 520, conteo);

                        linea = linea + 10;

                        gfx.DrawLine(pen, 15, linea, 600, linea);
                    }



                    //Fin Detalle Factura--------------------------------------------------------------------------------------------------------------------------------------------------------------------

                    var pen2 = new XPen(XColors.LightGray, 0.7);
                    pen.DashStyle = XDashStyle.Dot;

                    //Lineas Verticales del Detalle
                    gfx.DrawLine(pen2, 80, 130, 80, 550);
                    gfx.DrawLine(pen2, 310, 130, 310, 550);
                    gfx.DrawLine(pen2, 370, 130, 370, 550);
                    gfx.DrawLine(pen2, 500, 130, 500, 550);


                    //Pie de La Factura

                    gfx.DrawLine(pen, 15, 550, 600, 550);

                    gfx.DrawString("OBSERVACION", fontTitulo, pBrushes, 20, 560);
                    gfx.DrawString(dtEncabezado.Rows[0]["Observacion2"].ToString().Trim(), new XFont("Arial", 8), pBrushes, 20, 570);
                    gfx.DrawRectangle(XBrushes.LightGray, 420, 555, 180, 25);


                    //Total Factura


                    gfx.DrawString("TOTAL A PAGAR", new XFont("Arial", 9, XFontStyle.Bold), pBrushes, 430, 565);
                    gfx.DrawString(string.Format("$ {0:##,##}", Convert.ToInt64(dtEncabezado.Rows[0]["totalapagar"])), new XFont("Arial", 9, XFontStyle.Bold), pBrushes, 550, 565);
                    gfx.DrawString("ESTA FACTURA INCLUYE IVA POR", new XFont("Arial", 6), pBrushes, 430, 575);
                    gfx.DrawString(string.Format("$ {0:##,##}", Convert.ToInt64(dtEncabezado.Rows[0]["totaliva"])), new XFont("Arial", 8, XFontStyle.Bold), pBrushes, 550, 575);


                    var pen1 = new XPen(XColors.Black, 0);
                    pen1.DashStyle = XDashStyle.Solid;

                    gfx.DrawRectangle(pen1, 15, 580, 585, 20);
                    gfx.DrawString(dtEncabezado.Rows[0]["Observacion"].ToString().Trim(), new XFont("Arial", 6), pBrushes, 20, 587);
                    gfx.DrawString(dtEncabezado.Rows[0]["resolucion"].ToString().Trim(), new XFont("Arial", 6), pBrushes, 20, 597);

                    //Fin Pie de La FActura


                    //COLILLA DE PAGO 

                    var fontColilla = new XFont("Arial", 8, XFontStyle.Bold);
                    gfx.DrawImage(logo, 15, 605);

                    gfx.DrawString(dtEncabezado.Rows[0]["nombreempresa"].ToString().Trim(), new XFont("Arial", 8, XFontStyle.Bold), pBrushes, 80, 615);
                    gfx.DrawString(string.Format("Nit:{0}", dtEncabezado.Rows[0]["nitempresa"]).ToString().Trim(), new XFont("Arial", 8, XFontStyle.Bold), pBrushes, 80, 625);
                    gfx.DrawString(dtEncabezado.Rows[0]["ventadirecta"].ToString().Trim(), new XFont("Arial", 8, XFontStyle.Bold), pBrushes, 80, 635);
                    gfx.DrawString("CUENTAS BANCARIAS", new XFont("Arial", 8, XFontStyle.Bold), pBrushes, 20, 655);

                    int conteoBanco = 655;
                    for (int i = 0; i < dtBancos.Rows.Count; i++)
                    {
                        conteoBanco = conteoBanco + 10;
                        gfx.DrawString(dtBancos.Rows[i]["nombre"].ToString().Trim(), new XFont("Arial", 6), pBrushes, 20, conteoBanco);
                        conteoBanco = conteoBanco + 10;
                        gfx.DrawString(dtBancos.Rows[i]["descripcion"].ToString().Trim(), new XFont("Arial", 6), pBrushes, 20, conteoBanco);
                    }

                    gfx.DrawString("-COPIA BANCO-", fontColilla, pBrushes, 240, 615);

                    gfx.DrawRectangle(XBrushes.LightGray, 150, 655, 150, 30);

                    gfx.DrawString(dtEncabezado.Rows[0]["Asesora"].ToString().Trim(), new XFont("Arial", 5, XFontStyle.Bold), pBrushes, 152, 660);
                    gfx.DrawString(string.Format("CEDULA {0}", dtEncabezado.Rows[0]["cedula"].ToString().Trim()), new XFont("Arial", 5), pBrushes, 152, 665);
                    gfx.DrawString(string.Format("ZONA {0}", dtEncabezado.Rows[0]["Zona"].ToString().Trim()), new XFont("Arial", 5), pBrushes, 152, 670);

                    gfx.DrawString(string.Format("FECHA VENCIMIENTO {0}", Convert.ToDateTime(dtEncabezado.Rows[0]["fechavence"]).Date.ToShortDateString()), new XFont("Arial", 5, XFontStyle.Bold), pBrushes, 215, 665);
                    gfx.DrawString("VALOR:", new XFont("Arial", 10, XFontStyle.Bold), pBrushes, 200, 680);
                    gfx.DrawString(string.Format("$ {0:##,##}", Convert.ToInt64(dtEncabezado.Rows[0]["totalapagar"])), new XFont("Arial", 10, XFontStyle.Bold), pBrushes, 240, 680);


                    //Codigo de Barras
                    gfx.DrawString(dtEncabezado.Rows[0]["referenciabanco"].ToString().Trim(), new XFont("Arial", 5), pBrushes, 20, 725);
                    var xpoin = new XPoint(20, 730);
                    gfx.DrawBarCode(new Code3of9Standard(dtEncabezado.Rows[0]["codbarras"].ToString().Trim(), new XSize(250, 30)), pBrushes, new XFont("Arial", 10), xpoin);
                    gfx.DrawString(dtEncabezado.Rows[0]["codbarras2"].ToString().Trim(), new XFont("Arial", 5), pBrushes, 20, 765);

                    //Colilla Copia ---------------------------------------------------------------------------------------------------------------------------------


                    gfx.DrawImage(logo, 315, 605);

                    gfx.DrawString(dtEncabezado.Rows[0]["nombreempresa"].ToString().Trim(), new XFont("Arial", 8, XFontStyle.Bold), pBrushes, 380, 615);
                    gfx.DrawString(string.Format("Nit:{0}", dtEncabezado.Rows[0]["nitempresa"]).ToString().Trim(), new XFont("Arial", 8, XFontStyle.Bold), pBrushes, 380, 625);
                    gfx.DrawString(dtEncabezado.Rows[0]["ventadirecta"].ToString().Trim(), new XFont("Arial", 8, XFontStyle.Bold), pBrushes, 380, 635);
                    gfx.DrawString("CUENTAS BANCARIAS", new XFont("Arial", 8, XFontStyle.Bold), pBrushes, 320, 655);

                    int conteoBanco1 = 655;
                    for (int i = 0; i < dtBancos.Rows.Count; i++)
                    {
                        conteoBanco1 = conteoBanco1 + 10;
                        gfx.DrawString(dtBancos.Rows[i]["nombre"].ToString().Trim(), new XFont("Arial", 6), pBrushes, 320, conteoBanco1);
                        conteoBanco1 = conteoBanco1 + 10;
                        gfx.DrawString(dtBancos.Rows[i]["descripcion"].ToString().Trim(), new XFont("Arial", 6), pBrushes, 320, conteoBanco1);
                    }

                    gfx.DrawString("-COPIA DOLCE-", fontColilla, pBrushes, 540, 615);

                    gfx.DrawRectangle(XBrushes.LightGray, 450, 655, 150, 30);

                    gfx.DrawString(dtEncabezado.Rows[0]["Asesora"].ToString().Trim(), new XFont("Arial", 5, XFontStyle.Bold), pBrushes, 452, 660);
                    gfx.DrawString(string.Format("CEDULA {0}", dtEncabezado.Rows[0]["cedula"].ToString().Trim()), new XFont("Arial", 5), pBrushes, 452, 665);
                    gfx.DrawString(string.Format("ZONA {0}", dtEncabezado.Rows[0]["Zona"].ToString().Trim()), new XFont("Arial", 5), pBrushes, 452, 670);

                    gfx.DrawString(string.Format("FECHA VENCIMIENTO {0}", Convert.ToDateTime(dtEncabezado.Rows[0]["fechavence"]).Date.ToShortDateString()), new XFont("Arial", 5, XFontStyle.Bold), pBrushes, 515, 665);
                    gfx.DrawString("VALOR:", new XFont("Arial", 10, XFontStyle.Bold), pBrushes, 500, 680);
                    gfx.DrawString(string.Format("$ {0:##,##}", Convert.ToInt64(dtEncabezado.Rows[0]["totalapagar"])), new XFont("Arial", 10, XFontStyle.Bold), pBrushes, 540, 680);


                    //Codigo de Barras
                    gfx.DrawString(dtEncabezado.Rows[0]["referenciabanco"].ToString().Trim(), new XFont("Arial", 5), pBrushes, 320, 725);
                    var xpoin1 = new XPoint(320, 730);
                    gfx.DrawBarCode(new Code3of9Standard(dtEncabezado.Rows[0]["codbarras"].ToString().Trim(), new XSize(250, 30)), pBrushes, xpoin1);
                    gfx.DrawString(dtEncabezado.Rows[0]["codbarras2"].ToString().Trim(), new XFont("Arial", 5), pBrushes, 320, 765);


                    //Colilla Copia ---------------------------------------------------------------------------------------------------------------------------------

                    //FIN COLILLA DE PAGO

                    //FIN PAGINA FRONTAL

                    //pagina Posterior-------------------------------------------------------------------------------------------------------------------------------
                    PdfPage pgPosterior = _pdf.AddPage();
                    pgPosterior.Size = PdfSharp.PageSize.Letter;

                    XGraphics gfx1 = XGraphics.FromPdfPage(pgPosterior);

                    gfx1.DrawImage(logo, 220, 10);
                    gfx1.DrawString(dtEncabezado.Rows[0]["nombreempresa"].ToString().Trim(), new XFont("Arial", 8, XFontStyle.Bold), pBrushes, 275, 20);
                    gfx1.DrawString(string.Format("Nit:{0}", dtEncabezado.Rows[0]["nitempresa"]).ToString().Trim(), new XFont("Arial", 8, XFontStyle.Bold), pBrushes, 275, 30);
                    gfx1.DrawString(dtEncabezado.Rows[0]["ventadirecta"].ToString().Trim(), new XFont("Arial", 8, XFontStyle.Bold), pBrushes, 275, 40);

                    gfx1.DrawString(string.Format("FACTURA DE VENTA No {0}-{1}", dtEncabezado.Rows[0]["Prefijo"].ToString().Trim(), dtEncabezado.Rows[0]["numFactura"].ToString().Trim()), new XFont("Arial", 8, XFontStyle.Bold), pBrushes, 390, 20);
                    gfx1.DrawString(string.Format("PEDIDO No {0}", dtEncabezado.Rows[0]["numPedido"].ToString().Trim()), new XFont("Arial", 8, XFontStyle.Bold), pBrushes, 390, 30);

                    gfx1.DrawString("CAMPAÑA", new XFont("Arial", 8, XFontStyle.Bold), pBrushes, 550, 20);
                    gfx1.DrawString(dtEncabezado.Rows[0]["StrCampaña"].ToString().Trim(), new XFont("Arial", 8, XFontStyle.Bold), pBrushes, 550, 30);

                    gfx1.DrawRectangle(XBrushes.LightGray, 15, 15, 200, 30);


                    var listaTotalInformacion = dtInformativo.Select("strtipo=7");


                    gfx1.DrawString("PEDIDO No", new XFont("Arial", 5), pBrushes, 20, 20);
                    gfx1.DrawString("ZONA", new XFont("Arial", 5), pBrushes, 100, 20);
                    gfx1.DrawString("TOTAL PRENDAS", new XFont("Arial", 5), pBrushes, 150, 20);
                    gfx1.DrawString(listaTotalInformacion[0]["strObservacion"].ToString().Trim(), new XFont("Arial", 20, XFontStyle.Bold), pBrushes, 20, 40);
                    gfx1.DrawString(listaTotalInformacion[0]["strCodigo"].ToString().Trim(), new XFont("Arial", 20, XFontStyle.Bold), pBrushes, 90, 40);
                    gfx1.DrawString(listaTotalInformacion[0]["intCantidad"].ToString().Trim(), new XFont("Arial", 20, XFontStyle.Bold), pBrushes, 150, 40);

                    gfx1.DrawRectangle(XBrushes.LightGray, 15, 50, 585, 10);

                    gfx1.DrawString("INFORMATIVO DOLCE", new XFont("Arial", 9, XFontStyle.Bold), pBrushes, 20, 58);

                    var pen3 = new XPen(XColors.LightGray, 0.7);
                    pen.DashStyle = XDashStyle.Solid;

                    //Lineas Verticales del Detalle
                    gfx1.DrawLine(pen3, 350, 60, 350, 600);

                    //detalle de Premio
                    gfx1.DrawString("PREMIOS ENTREGADOS", new XFont("Arial", 8, XFontStyle.Bold), pBrushes, 20, 70);

                    int conteoPremios = 70;

                    for (int i = 0; i < dtPremios.Rows.Count; i++)
                    {
                        conteoPremios = conteoPremios + 10;
                        gfx1.DrawString(string.Format("{0} CANTIDAD {1}", dtPremios.Rows[i]["strCodigo"].ToString().Trim(), dtPremios.Rows[i]["intCantidad"]), new XFont("Arial", 7), pBrushes, 20, conteoPremios);
                    }

                    //Fin Detalle Premios 


                    //Recordatorios y Otros

                    var listaOtros = dtInformativo.Select("strtipo=3");

                    gfx1.DrawLine(pen3, 15, 300, 350, 300);
                    gfx1.DrawString("RECORDATORIOS Y OTROS", new XFont("Arial", 8, XFontStyle.Bold), pBrushes, 20, 310);

                    int conteoOtros = 310;

                    for (int i = 0; i < listaOtros.Count(); i++)
                    {
                        conteoOtros = conteoOtros + 10;

                        string testo = listaOtros[i]["strObservacion"].ToString().Trim();

                        if (testo.Substring(0, 3) == "[N]")
                        {
                            gfx1.DrawString(listaOtros[i]["strObservacion"].ToString().Trim().Substring(3), new XFont("Arial", 7, XFontStyle.Bold), pBrushes, 20, conteoOtros);
                        }
                        else
                        {
                            gfx1.DrawString(listaOtros[i]["strObservacion"].ToString().Trim(), new XFont("Arial", 5), pBrushes, 20, conteoOtros);
                        }

                    }

                    //Fin REcordatorios y Otros


                    //Productos Agotados

                    gfx1.DrawString("PRODUCTOS AGOTADOS", new XFont("Arial", 8, XFontStyle.Bold), pBrushes, 355, 70);
                    gfx1.DrawRectangle(XBrushes.LightGray, 350, 75, 250, 10);
                    gfx1.DrawString("CODIGO", new XFont("Arial", 5), pBrushes, 355, 82);
                    gfx1.DrawString("DESCRIPCION", new XFont("Arial", 5), pBrushes, 450, 82);
                    gfx1.DrawString("CANTIDAD", new XFont("Arial", 5), pBrushes, 570, 82);

                    gfx1.DrawLine(pen2, 400, 85, 400, 330);
                    gfx1.DrawLine(pen2, 560, 85, 560, 330);

                    var listaAgotados = dtInformativo.Select("strtipo=4");

                    if (listaAgotados.Count() > 0)
                    {
                        int conteoAgotados = 82;

                        for (int i = 0; i < listaAgotados.Count(); i++)
                        {

                            conteoAgotados = conteoAgotados + 10;
                            gfx1.DrawString(listaAgotados[i]["strCodigo"].ToString().Trim(), new XFont("Arial", 6), pBrushes, 355, conteoAgotados);
                            gfx1.DrawString(listaAgotados[i]["strObservacion"].ToString().Trim(), new XFont("Arial", 6), pBrushes, 405, conteoAgotados);
                            gfx1.DrawString(listaAgotados[i]["intCantidad"].ToString().Trim(), new XFont("Arial", 6), pBrushes, 570, conteoAgotados);

                        }
                    }





                    //Fin Productos Agotados

                    //Cambios Surtidos

                    gfx1.DrawLine(pen3, 350, 330, 600, 330);

                    gfx1.DrawString("CAMBIOS SURTIDOS", new XFont("Arial", 8, XFontStyle.Bold), pBrushes, 355, 340);
                    gfx1.DrawRectangle(XBrushes.LightGray, 350, 345, 250, 10);
                    gfx1.DrawString("CODIGO", new XFont("Arial", 5), pBrushes, 355, 352);
                    gfx1.DrawString("DESCRIPCION", new XFont("Arial", 5), pBrushes, 450, 352);
                    gfx1.DrawString("CANTIDAD", new XFont("Arial", 5), pBrushes, 570, 352);

                    gfx1.DrawLine(pen2, 400, 350, 400, 480);
                    gfx1.DrawLine(pen2, 560, 350, 560, 480);


                    var listaCambios = dtInformativo.Select("strtipo=5");

                    if (listaCambios.Count() > 0)
                    {

                        var conteoCambios = 352;

                        for (int i = 0; i < listaCambios.Count(); i++)
                        {
                            conteoCambios = conteoCambios + 10;
                            gfx1.DrawString(listaCambios[i]["strCodigo"].ToString().Trim(), new XFont("Arial", 6), pBrushes, 355, conteoCambios);
                            gfx1.DrawString(listaCambios[i]["strObservacion"].ToString().Trim(), new XFont("Arial", 6), pBrushes, 405, conteoCambios);
                            gfx1.DrawString(listaCambios[i]["intCantidad"].ToString().Trim(), new XFont("Arial", 6), pBrushes, 570, conteoCambios);
                        }
                    }


                    //Fin Cambios Surtidos


                    //Cambios Agotados

                    gfx1.DrawLine(pen3, 350, 480, 600, 480);

                    gfx1.DrawString("CAMBIOS AGOTADOS", new XFont("Arial", 8, XFontStyle.Bold), pBrushes, 355, 490);
                    gfx1.DrawRectangle(XBrushes.LightGray, 350, 495, 250, 10);
                    gfx1.DrawString("CODIGO", new XFont("Arial", 5), pBrushes, 355, 502);
                    gfx1.DrawString("DESCRIPCION", new XFont("Arial", 5), pBrushes, 450, 502);
                    gfx1.DrawString("CANTIDAD", new XFont("Arial", 5), pBrushes, 570, 502);

                    gfx1.DrawLine(pen2, 400, 500, 400, 600);
                    gfx1.DrawLine(pen2, 560, 500, 560, 600);

                    gfx1.DrawLine(pen2, 15, 600, 600, 600);


                    var listaCambiosAgotados = dtInformativo.Select("strtipo=6");


                    if (listaCambiosAgotados.Count() > 0)
                    {
                        int conteoCambiosAgotados = 502;

                        for (int i = 0; i < listaCambiosAgotados.Count(); i++)
                        {
                            conteoCambiosAgotados = conteoCambiosAgotados + 10;
                            gfx1.DrawString(listaCambiosAgotados[i]["strCodigo"].ToString().Trim(), new XFont("Arial", 6), pBrushes, 355, conteoCambiosAgotados);
                            gfx1.DrawString(listaCambiosAgotados[i]["strObservacion"].ToString().Trim(), new XFont("Arial", 6), pBrushes, 405, conteoCambiosAgotados);
                            gfx1.DrawString(listaCambiosAgotados[i]["intCantidad"].ToString().Trim(), new XFont("Arial", 6), pBrushes, 570, conteoCambiosAgotados);
                        }
                    }


                    //Fin Cambios Agotados

                    DataTable dtPuntos = dts.Tables[5];

                    if (dtPuntos.Rows.Count > 0)
                    {
                        Int32 Acumulados = Convert.ToInt32(dtPuntos.Rows[0]["Puntos_Acumulados"]);
                        Int32 total = Convert.ToInt32(dtPuntos.Rows[0]["Saldo"]);
                        Int32 devoluciones = Convert.ToInt32(dtPuntos.Rows[0]["Devoluciones"]);
                        Int32 utilizados = Convert.ToInt32(dtPuntos.Rows[0]["Puntos_Utilizados"]);
                        Int32 pendientes = Convert.ToInt32(dtPuntos.Rows[0]["Puntos_Pendientes"]);

                        int saldo = ((Acumulados - utilizados) - devoluciones);


                        XImage anitaDolce = Properties.Resources.Anita;

                        gfx1.DrawImage(anitaDolce,20,650);

                        gfx1.DrawString(string.Format("PUNTOS ACUMULADOS A CAMPAÑA {0}", dtPuntos.Rows[0]["Campaña_anterior"]),new XFont("Tahoma",12,XFontStyle.Bold),pBrushes,200,650);
                        gfx1.DrawString("PUNTOS", new XFont("Tahoma",10,XFontStyle.Bold),pBrushes,200,680);
                        gfx1.DrawString("PUNTOS UTILIZADOS", new XFont("Tahoma", 10, XFontStyle.Bold), pBrushes, 270, 680);
                        gfx1.DrawString("DEVOLUCIONES", new XFont("Tahoma", 10, XFontStyle.Bold), pBrushes, 400, 680);
                        gfx1.DrawString("TOTALES", new XFont("Tahoma", 10, XFontStyle.Bold), pBrushes, 500, 680);


                        gfx1.DrawString(dtPuntos.Rows[0]["Puntos_Acumulados"].ToString(),new XFont("Tahoma",9,XFontStyle.Bold),pBrushes,200,710);
                        gfx1.DrawString(dtPuntos.Rows[0]["Puntos_Utilizados"].ToString(), new XFont("Tahoma", 9, XFontStyle.Bold), pBrushes, 270, 710);
                        gfx1.DrawString(dtPuntos.Rows[0]["Devoluciones"].ToString(), new XFont("Tahoma", 9, XFontStyle.Bold), pBrushes, 400, 710);
                        gfx1.DrawString(saldo.ToString(), new XFont("Tahoma", 9, XFontStyle.Bold), pBrushes, 500, 710);


                        gfx1.DrawLine(pen,200,720,600,720);

                        gfx1.DrawString("PUNTOS ACUMULADOS EN ESTA CAMPAÑA", new XFont("Tahoma", 9, XFontStyle.Bold), pBrushes, 200, 740);
                        gfx1.DrawString(pendientes.ToString(), new XFont("Tahoma", 16, XFontStyle.Bold), pBrushes, 450, 740);
                        gfx1.DrawString("SOLO SI HACES TU PAGO A TIEMPO",new XFont("Tahoma", 9, XFontStyle.Bold), pBrushes,200,760);
                        gfx1.DrawString("¡EXITOS Y A REDIMIR MUCHOS PREMIOS!", new XFont("Segoe Script", 14), pBrushes, 200, 780);
                        
                    }


                    //gfx1.DrawString("RESPALDO COLILLA DE PAGO", new XFont("Arial", 15), XBrushes.LightGray, 20, 700);
                    //gfx1.DrawString("RESPALDO COLILLA DE PAGO", new XFont("Arial", 15), XBrushes.LightGray, 350, 700);

                    //Fin pagina Posterior-------------------------------------------------------------------------------------------------------------------------------
                    //Se Crea El Archivo Pdf


                    strMensaje = "";
                    logRespuesta = true;
                    
                    //string carpeta = Properties.Settings.Default.Ruta_Pdf.ToString();
                    string xarchivo = string.Format("{0}-{1}.pdf", dtEncabezado.Rows[0]["Prefijo"].ToString().Trim(),
                        dtEncabezado.Rows[0]["numFactura"].ToString().Trim());
                    string strNombrepdf = System.IO.Path.Combine(path, xarchivo); 
                        //string.Format("{0}{1}", carpeta,xarchivo);
                    

                    strArchivo = xarchivo;

                    _pdf.Save(strNombrepdf);

                   // Process.Start(strNombrepdf);


                }
                else
                {
                    strMensaje = "";
                    logRespuesta = true;
                    strArchivo = "";
                }
            }
            else
            {
                strMensaje = "";
                logRespuesta = true;
                strArchivo = "";
            }

        }
    }
}
