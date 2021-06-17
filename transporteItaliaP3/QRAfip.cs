using System;
using System.Collections.Generic;
using System.Text;

namespace transporteItaliaP3
{
    class QRAfip
    {
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
