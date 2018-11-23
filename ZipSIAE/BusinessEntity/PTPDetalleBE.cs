using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessEntity
{
    public class PTPDetalleBE
    {
        public PTPBE PTP { get; set; }
        public NodoBE NodoB { get; set; }
        public Double AzimuthNodoA { get; set; }
        public Double AzimuthNodoB { get; set; }
        public Double ElevacionNodoA { get; set; }
        public Double ElevacionNodoB { get; set; }
        public Double Distancia { get; set; }
        public Double CotaNodoB { get; set; }
        public String ModeloAntenaNodoA { get; set; }
        public String ModeloAntenaNodoB { get; set; }
        public String DiametroAntenaNodoA { get; set; }
        public String DiametroAntenaNodoB { get; set; }
        public Int32 AlturaAntenaNodoA { get; set; }
        public Int32 AlturaAntenaNodoB { get; set; }
        public Double GananciaAntenaNodoA { get; set; }
        public Double GananciaAntenaNodoB { get; set; }
        public String IdCanalNodoA { get; set; }
        public String IdCanalNodoB { get; set; }
        public String DisenoFrecuenciaNodoA { get; set; }
        public String DisenoFrecuenciaNodoB { get; set; }
        public EntidadDetalleBE Polarizacion { get; set; }
        public String ModeloRadioNodoA { get; set; }
        public String DesignadorEmisionNodoA { get; set; }
        public Int32 PotenciaTorreNodoA { get; set; }
        public Int32 PotenciaTorreNodoB { get; set; }
        public Double EIRPNodoA { get; set; }
        public Double EIRPNodoB { get; set; }
        public Double NivelUmbralNodoA { get; set; }
        public Double SenalRecepcionNodoA { get; set; }
        public Double SenalRecepcionNodoB { get; set; }
        public Double MargenEfectividadDesvanecimientoNodoA { get; set; }
        public Double MargenEfectividadDesvanecimientoNodoB { get; set; }
        public Double DisponibilidadAnualMultirutasNodoA { get; set; }
        public Double DisponibilidadAnualLluviaNodoA { get; set; }
        public Double DisponibilidadAnualMultirutasLluvia { get; set; }

        public PTPDetalleBE()
        {
            PTP = new PTPBE();
            NodoB = new NodoBE();
            Polarizacion = new EntidadDetalleBE();
        }
    }
}
