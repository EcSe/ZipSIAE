using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessEntity
{
    
    public class EntidadDetalleBE
    {
        public EntidadBE Entidad { get; set; }
        public String IdValor { get; set; }
        public String ValorCadena1 { get; set; }
        public String ValorCadena2 { get; set; }
        public String ValorCadena3 { get; set; }
        public String ValorCadena4 { get; set; }
        public String ValorCadena5 { get; set; }
        public Int32 ValorEntero1 { get; set; }
        public Int32 ValorEntero2 { get; set; }
        public Int32 ValorEntero3 { get; set; }
        public Int32 ValorEntero4 { get; set; }
        public Int32 ValorEntero5 { get; set; }
        public Double ValorNumerico1 { get; set; }
        public Double ValorNumerico2 { get; set; }
        public Double ValorNumerico3 { get; set; }
        public Double ValorNumerico4 { get; set; }
        public Double ValorNumerico5 { get; set; }
        public DateTime ValorFecha1 { get; set; }
        public DateTime ValorFecha2 { get; set; }
        public DateTime ValorFecha3 { get; set; }
        public DateTime ValorFecha4 { get; set; }
        public DateTime ValorFecha5 { get; set; }
        public Boolean ValorBooleano1 { get; set; }
        public Boolean ValorBooleano2 { get; set; }
        public Boolean ValorBooleano3 { get; set; }
        public Boolean ValorBooleano4 { get; set; }
        public Boolean ValorBooleano5 { get; set; }
        public Byte[] ValorBinario1 { get; set; }
        public Byte[] ValorBinario2 { get; set; }
        public Byte[] ValorBinario3 { get; set; }
        public Byte[] ValorBinario4 { get; set; }
        public Byte[] ValorBinario5 { get; set; }
        public String Metodo { get; set; }
        public EntidadDetalleBE EntidadDetalleSecundario { get; set; }
        public UsuarioBE UsuarioCreacion { get; set; }

        public EntidadDetalleBE()
        {
            Entidad = new EntidadBE();
            ValorCadena1 = string.Empty;
            ValorCadena2 = string.Empty;
            ValorCadena3 = string.Empty;
            ValorCadena4 = string.Empty;
            ValorCadena5 = string.Empty;
            IdValor = String.Empty;
            Metodo = String.Empty;
        }

        public EntidadDetalleBE Clone()
        {
            return (EntidadDetalleBE)this.MemberwiseClone();
        }
    }
}
