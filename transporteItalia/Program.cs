using System;
using System.Collections.Generic;
using System.IO;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System.Drawing.Printing;
using System.Diagnostics;
using Newtonsoft.Json;
using ZXing;
using System.Drawing;
using System.Globalization;

namespace transporteItalia
{
    //TODO acomodar todo con el cuerpo mal, asi se manda y asi para corregir
            //ver para imprimir el pdf desde una libreria
            //qr ver los margenes



    class Program
    {

        // Declaracion de fuentes
        public static XFont fontCourier10 = new XFont("Courier New", 10, XFontStyle.Regular);
        public static XFont fontCourier8 = new XFont("Courier New", 8, XFontStyle.Regular);
        public static XFont fontCourier7 = new XFont("Courier New", 7, XFontStyle.Regular);
        public static XFont fontCourier6 = new XFont("Courier New", 6, XFontStyle.Regular);
        public static XFont fontCourierBold35 = new XFont("Courier New", 35, XFontStyle.Bold);
        public static XFont fontCourierBold20 = new XFont("Courier New", 20, XFontStyle.Bold);
        public static XFont fontCourierBold15 = new XFont("Courier New", 15, XFontStyle.Bold);
        public static XFont fontCourierBold14 = new XFont("Courier New", 14, XFontStyle.Bold);
        public static XFont fontCourierBold13 = new XFont("Courier New", 13, XFontStyle.Bold);
        public static XFont fontCourierBold10 = new XFont("Courier New", 10, XFontStyle.Bold);
        public static XFont fontCourierBold9 = new XFont("Courier New", 9, XFontStyle.Bold);
        public static XFont fontCourierBold7 = new XFont("Courier New", 7, XFontStyle.Bold);
        public static XFont fontHelvetica35 = new XFont("Helvetica", 35, XFontStyle.Bold);
        /////////////////////////
        static void Main(string[] args)
        {
            string path = Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "IN_FACTU");
            string textToParse = System.IO.File.ReadAllText(path);
            //int paginas = textToParse.Length / 6615;

            //Creamos un documento unico
            PdfDocument document = new PdfDocument();
            document.Info.Title = "Transporte Italia - Arrecifes";

            document.Options.FlateEncodeMode = PdfFlateEncodeMode.BestCompression;
            pdfGeneratorTransItalia(textToParse, document);
            string filename = "pdfTransporteItalia.pdf";

            document.Save(filename);

            // File.Delete(filename);
            System.Diagnostics.Process.Start(filename);
            string cPrinter = GetDefaultPrinter();
            string cRun = "PDFlite.exe";
            string arguments = " -print-to \"" + cPrinter + "\" " + " -print-settings \"" + "1x" + "\" " + filename;
            string argument = "-print-to-default " + filename;
            Process process = new Process();
            //process.StartInfo.FileName = cRun;
            // process.StartInfo.Arguments = argument;
            //process.Start();
            //  process.WaitForExit();

            // File.Delete(filename);
        }

        static void testImpresion()
        {
            PrintDocument test = new System.Drawing.Printing.PrintDocument();


        }


        static string GetDefaultPrinter()
        {
            PrinterSettings settings = new PrinterSettings();
            foreach (string printer in PrinterSettings.InstalledPrinters)
            {
                settings.PrinterName = printer;
                if (settings.IsDefaultPrinter)
                    return printer;
            }
            return string.Empty;
        }

        private static void pdfGeneratorTransItalia(string pagina, PdfDocument document)
        {
            int pivote = 0;
            PdfPage page = document.AddPage();
            page.Size = PdfSharp.PageSize.A4;

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);
            XImage img = XImage.FromFile("TransporteItalia.jpg");
            gfx.DrawImage(img, 0, 0);

            //Armado de variables
            String tipoComprobante = pagina.Substring(pivote, 3);
            String letra = pagina.Substring(pivote += 3, 1);
            String prefijo = pagina.Substring(pivote += 1, 4);
            String numero = pagina.Substring(pivote += 4, 8);

            String fecha = pagina.Substring(pivote += 8, 2) + "/" + pagina.Substring(pivote += 2, 2) + "/" + pagina.Substring(pivote += 2, 4);
            //
            String ma_cuenta = pagina.Substring(pivote += 4, 8);
            String ma_nombre = pagina.Substring(pivote += 8, 30);
            String ma_domicilio = pagina.Substring(pivote += 30, 30);
            String ma_localidad = pagina.Substring(pivote += 30, 30);

            String ma_cuit = pagina.Substring(pivote += 30, 11);
            String ma_condIva = pagina.Substring(pivote += 11, 15);

            List<string> cuerpos = new List<string>();//Tabla Cuerpo
            cuerpos.Add(pagina.Substring(pivote += 15, 80));
            for (int i = 0; i < 11; i++) cuerpos.Add(pagina.Substring(pivote += 80, 80));

            String subtotal = int.Parse(pagina.Substring(pivote += 80, 12)).ToString();
            subtotal = subtotal.Insert(subtotal.Length - 2, ",");
            String iva = int.Parse(pagina.Substring(pivote += 12, 12)).ToString();
            iva = iva.Insert(iva.Length - 2, ",");
            String ivaNoInscripto = int.Parse(pagina.Substring(pivote += 12, 12)).ToString();
            if (ivaNoInscripto == "0") ivaNoInscripto = "";
            else ivaNoInscripto = ivaNoInscripto.Insert(ivaNoInscripto.Length - 2, ",");
            String exento = pagina.Substring(pivote += 12, 12);
            String total = int.Parse(pagina.Substring(pivote += 12, 12)).ToString();
            total = total.Insert(total.Length - 2, ",");

            String loDaVuelta = pagina.Substring(pivote += 12, 1);
            
            String re_cuenta = pagina.Substring(pivote += 1, 8);
            String re_nombre = pagina.Substring(pivote += 8, 30);
            String re_domicilio = pagina.Substring(pivote += 30, 30);
            String re_localidad = pagina.Substring(pivote += 30, 30);

            String re_cuit = pagina.Substring(pivote += 30, 11);
            String re_condIva = pagina.Substring(pivote += 11, 15);

            String redespacho = pagina.Substring(pivote += 15, 30);
            String cae = pagina.Substring(pivote += 30, 14);
            String caeVto = pagina.Substring(pivote += 14, 4) + "/" + pagina.Substring(pivote += 4, 2) + "/" + pagina.Substring(pivote += 2, 2);

            gfx.DrawString(letra, fontHelvetica35, XBrushes.Black, 285, 41);
            gfx.DrawString("Código:" + tipoComprobante, fontCourier8, XBrushes.Black, 275, 54);

            gfx.DrawString(letra, fontHelvetica35, XBrushes.Black, 285, 458);
            gfx.DrawString("Código:" + tipoComprobante, fontCourier8, XBrushes.Black, 275, 471);

            gfx.DrawString("Número: ", fontCourierBold10, XBrushes.Black, 386, 25);
            gfx.DrawString(prefijo + "," + numero, fontCourierBold10, XBrushes.Black, 440, 25);
            gfx.DrawString("Fecha: ", fontCourierBold10, XBrushes.Black, 386, 40);
            gfx.DrawString(fecha, fontCourierBold10, XBrushes.Black, 440, 40);

            gfx.DrawString("Número: ", fontCourierBold10, XBrushes.Black, 386, 443);
            gfx.DrawString(prefijo + "," + numero, fontCourierBold10, XBrushes.Black, 440, 443);
            gfx.DrawString("Fecha: ", fontCourierBold10, XBrushes.Black, 386, 458);
            gfx.DrawString(fecha, fontCourierBold10, XBrushes.Black, 440, 458);


            if (loDaVuelta == "N" || loDaVuelta == "n")
            {
                drawNomDomLoc(gfx, re_nombre, re_domicilio, re_localidad);
                drawNomDomLocIvaCuit(gfx, ma_nombre, ma_domicilio, ma_localidad, ma_condIva, ma_cuit);
                DrawQR(gfx, fecha,ma_cuit, Int32.Parse(prefijo), Int32.Parse(tipoComprobante), Int32.Parse(numero), Double.Parse(total),cae);
            }
            else
            {
                drawNomDomLoc(gfx, ma_nombre, ma_domicilio, ma_localidad);
                drawNomDomLocIvaCuit(gfx, re_nombre, re_domicilio, re_localidad, re_condIva, re_cuit);
                DrawQR(gfx, fecha, re_cuit, Int32.Parse(prefijo), Int32.Parse(tipoComprobante), Int32.Parse(numero), Double.Parse(total), cae);
            }

            gfx.DrawString("REDESPACHAR POR: ", fontCourier10, XBrushes.Black, 26, 385);
            gfx.DrawString(redespacho, fontCourier10, XBrushes.Black, 140, 385);
            gfx.DrawString("REDESPACHAR POR: ", fontCourier10, XBrushes.Black, 26, 800);
            gfx.DrawString(redespacho, fontCourier10, XBrushes.Black, 140, 800);


            int posy = 340;
            gfx.DrawString("SUB-TOTAL: ", fontCourierBold10, XBrushes.Black, 370, posy += 15);
            gfx.DrawString("$ " + subtotal, fontCourier10, XBrushes.Black, 470, posy);
            gfx.DrawString("IVA 21%: ", fontCourier10, XBrushes.Black, 370, posy += 15);
            gfx.DrawString("$ " + iva, fontCourier10, XBrushes.Black, 470, posy);
            gfx.DrawString("IVA NO I.: ", fontCourier10, XBrushes.Black, 370, posy += 15);
            gfx.DrawString("$ " + ivaNoInscripto, fontCourier10, XBrushes.Black, 470, posy);
            gfx.DrawString("TOTAL $: ", fontCourierBold10, XBrushes.Black, 370, posy += 15);
            gfx.DrawString("$ " + total, fontCourier10, XBrushes.Black, 470, posy);

            posy = 757;
            gfx.DrawString("SUB-TOTAL: ", fontCourierBold10, XBrushes.Black, 370, posy += 15);
            gfx.DrawString("$ " + subtotal, fontCourier10, XBrushes.Black, 470, posy);
            gfx.DrawString("IVA 21%: ", fontCourier10, XBrushes.Black, 370, posy += 15);
            gfx.DrawString("$ " + iva, fontCourier10, XBrushes.Black, 470, posy);
            gfx.DrawString("IVA NO I.: ", fontCourier10, XBrushes.Black, 370, posy += 15);
            gfx.DrawString("$ " + ivaNoInscripto, fontCourier10, XBrushes.Black, 470, posy);
            gfx.DrawString("TOTAL: ", fontCourierBold10, XBrushes.Black, 370, posy += 15);
            gfx.DrawString("$ " + total, fontCourier10, XBrushes.Black, 470, posy);

            gfx.DrawString("C.A.E.: " + cae + " Vencimiento: " + caeVto, fontCourierBold9, XBrushes.Black, 30, 403);
            gfx.DrawString("C.A.E.: " + cae + " Vencimiento: " + caeVto, fontCourierBold9, XBrushes.Black, 30, 820);
        }

        private static void drawNomDomLoc(XGraphics gfx, String nombre, String domicilio, String localidad)
        {
            int posy = 325;
            gfx.DrawString("DESTINATARIO: ", fontCourier10, XBrushes.Black, 26, posy += 15);
            gfx.DrawString(nombre, fontCourier10, XBrushes.Black, 140, posy);
            gfx.DrawString("DOMICILIO: ", fontCourier10, XBrushes.Black, 26, posy += 15);
            gfx.DrawString(domicilio, fontCourier10, XBrushes.Black, 140, posy);
            gfx.DrawString("LOCALIDAD: ", fontCourier10, XBrushes.Black, 26, posy += 15);
            gfx.DrawString(localidad, fontCourier10, XBrushes.Black, 140, posy);
            posy = 740;
            gfx.DrawString("DESTINATARIO: ", fontCourier10, XBrushes.Black, 26, posy += 15);
            gfx.DrawString(nombre, fontCourier10, XBrushes.Black, 140, posy);
            gfx.DrawString("DOMICILIO: ", fontCourier10, XBrushes.Black, 26, posy += 15);
            gfx.DrawString(domicilio, fontCourier10, XBrushes.Black, 140, posy);
            gfx.DrawString("LOCALIDAD: ", fontCourier10, XBrushes.Black, 26, posy += 15);
            gfx.DrawString(localidad, fontCourier10, XBrushes.Black, 140, posy);
        }

        private static void drawNomDomLocIvaCuit(XGraphics gfx, String nombre, String domicilio, String localidad, String conIva, String cuit)
        {
            int posy = 101;
            gfx.DrawString(nombre, fontCourierBold10, XBrushes.Black, 26, posy+=11);
            //gfx.DrawString("DESTINATARIO: ", fontCourier10, XBrushes.Black, 26, posy += 11);
            //gfx.DrawString("DOMICILIO: ", fontCourier10, XBrushes.Black, 26, posy += 11);
            gfx.DrawString(domicilio, fontCourierBold10, XBrushes.Black, 26, posy += 11);
            //gfx.DrawString("LOCALIDAD: ", fontCourier10, XBrushes.Black, 26, posy += 11);
            gfx.DrawString(localidad, fontCourierBold10, XBrushes.Black, 26, posy += 11);
            gfx.DrawString("COND. IVA: ", fontCourierBold10, XBrushes.Black, 250, posy-=6);
            gfx.DrawString(conIva, fontCourierBold10, XBrushes.Black, 360, posy);
            gfx.DrawString("CUIT: ", fontCourierBold10, XBrushes.Black, 250, posy -= 11);
            gfx.DrawString(cuit.Substring(0,2) + "-" + cuit.Substring(2,8) + "-" + cuit.Substring(10,1), fontCourierBold10, XBrushes.Black, 360, posy);
            posy = 518;
            //gfx.DrawString("DESTINATARIO: ", fontCourier10, XBrushes.Black, 26, posy += 11);
            gfx.DrawString(nombre, fontCourierBold10, XBrushes.Black, 26, posy+=11);
            //gfx.DrawString("DOMICILIO: ", fontCourier10, XBrushes.Black, 26, posy += 11);
            gfx.DrawString(domicilio, fontCourierBold10, XBrushes.Black, 26, posy+=11);
            //gfx.DrawString("LOCALIDAD: ", fontCourier10, XBrushes.Black, 26, posy += 11);
            gfx.DrawString(localidad, fontCourierBold10, XBrushes.Black, 26, posy+=11);
            gfx.DrawString("COND. IVA: ", fontCourierBold10, XBrushes.Black, 250, posy-=6);
            gfx.DrawString(conIva, fontCourierBold10, XBrushes.Black, 360, posy);
            gfx.DrawString("CUIT: ", fontCourierBold10, XBrushes.Black, 250, posy -=11);
            gfx.DrawString(cuit.Substring(0, 2) + "-" + cuit.Substring(2, 8) + "-" + cuit.Substring(10, 1), fontCourierBold10, XBrushes.Black, 360, posy);
        }

        private static void DrawQR(XGraphics gfx, String fecha, String cuit, int prefijo, int tipoComp, int numero, Double importe, String cae)
        {
            //Preparo la fecha
            DateTime d = DateTime.ParseExact(fecha, "dd/MM/yyyy", CultureInfo.InvariantCulture);
            var fechaParseada = d.ToString("yyyy-MM-dd");
            //Inicio el string que va a ir en el QR
            String qr = "https://www.afip.gob.ar/fe/qr/?p=";
            //Creo un objeto del tipo QRAfip con todo lo que me piden los delincuentes
            QRAfip qrAFIP = new QRAfip(
                1,//version
                fechaParseada, //fecha
                cuit,//cuit ma o re
                prefijo, //pto de venta
                tipoComp,
                numero,
                importe,//cambiar coma por punto
                "PES",
                "1",
                "E",
                cae
                );
            //Creo el JSON
            string jsonQRAfip = JsonConvert.SerializeObject(qrAFIP);
            //Lo paso a base64
            var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(jsonQRAfip);
            var jsonBase64 = System.Convert.ToBase64String(plainTextBytes);
            //Se lo agrego al qr para completarlo
            qr += jsonBase64;

            //Con esta prueba vemos si el JSON se paso a base64 Correctamente.
            /*var base64EncodedBytes = System.Convert.FromBase64String(jsonBase64);
            var pruebadecode = System.Text.Encoding.UTF8.GetString(base64EncodedBytes);*/

            //Draw QR1
            var bcWriter = new BarcodeWriter
            {
                Format = BarcodeFormat.QR_CODE,
                Options = new ZXing.Common.EncodingOptions
                {
                    Height = 125,
                    Width = 125,
                    PureBarcode = true,
                    Margin = 0
             
                },
            };
            Bitmap bm = bcWriter.Write(qr);
            XImage img = XImage.FromGdiPlusImage((Image)bm);
            img.Interpolate = false;
            gfx.DrawImage(img, 497, 88);


        }
        class QRAfip
        {
            Int32 ver;
            String fecha;
            String cuit;
            Int32 ptoVta;
            Int32 tipoComp;
            Int32 nroCmp;
            Double importe;
            String moneda;
            String ctz;
            // Int32 tipoDocRec;
            //Double nroDocRec;
            String tipoCodAut;
            String codAut;

            //Constructor
            public QRAfip(Int32 ver,
                String fecha,
                String cuit,
                Int32 ptoVta,
                Int32 tipoComp,
                Int32 nroCmp,
                Double importe,
                String moneda,
                String ctz,
                //Int32 tipoDocRec,
                //Double nroDocRec,
                String tipoCodAut,
                String codAut
                )
            {
                this.Ver = ver;
                this.Fecha = fecha;
                this.Cuit = cuit;
                this.PtoVta = ptoVta;
                this.TipoComp = tipoComp;
                this.NroCmp = nroCmp;
                this.Importe = importe;
                this.Moneda = moneda;
                this.Ctz = ctz;
                //this.tipoDocRec=tipoDocRec
                //this.nroDocRec =nroDocRec 
                this.TipoCodAut = tipoCodAut;
                this.CodAut = codAut;
            }

            public int Ver { get => ver; set => ver = value; }
            public string Fecha { get => fecha; set => fecha = value; }
            public string Cuit { get => cuit; set => cuit = value; }
            public int PtoVta { get => ptoVta; set => ptoVta = value; }
            public int TipoComp { get => tipoComp; set => tipoComp = value; }
            public int NroCmp { get => nroCmp; set => nroCmp = value; }
            public double Importe { get => importe; set => importe = value; }
            public string Moneda { get => moneda; set => moneda = value; }
            public string Ctz { get => ctz; set => ctz = value; }
            public string TipoCodAut { get => tipoCodAut; set => tipoCodAut = value; }
            public string CodAut { get => codAut; set => codAut = value; }
        }
    }
}