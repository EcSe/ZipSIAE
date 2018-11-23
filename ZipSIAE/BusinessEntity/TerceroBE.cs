using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessEntity
{
    public class TerceroBE
    {
        public EntidadDetalleBE TipoDocumento { get; set; }
        public String NumeroDocumento { get; set; }
        public String NombreRazon { get; set; }
        public String ApellidoPaterno { get; set; }
        public String ApellidoMaterno { get; set; }
        public String NombreCompleto { get; set; }
        public EntidadDetalleBE Actividad { get; set; }

        public TerceroBE()
        {
            TipoDocumento = new EntidadDetalleBE();
            NumeroDocumento = string.Empty;
            Actividad = new EntidadDetalleBE();
        }
    }
}
