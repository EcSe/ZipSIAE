using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessEntity
{
    public class PMPBE
    {
        public NodoBE Nodo { get; set; }
        public EntidadDetalleBE ModeloAntenaNodo { get; set; }
        public Int32 GananciaAntenaNodo { get; set; }
        public Int32 AlturaAntenaNodo { get; set; }
        public Int32 AzimuthAntenaNodo { get; set; }
        public Int32 ElevacionAntenaNodo { get; set; }
        public Int32 EIRPAntenaNodo { get; set; }
        public UsuarioBE UsuarioCreacion { get; set; }
        public UsuarioBE UsuarioModificacion { get; set; }
        public List<PMPDetalleBE> Detalles { get; set; }
        public PMPBE()
        {
            Nodo = new NodoBE();
            ModeloAntenaNodo = new EntidadDetalleBE();
            Detalles = new List<PMPDetalleBE>();
            UsuarioCreacion = new UsuarioBE();
            UsuarioModificacion = new UsuarioBE();
        }
    }
}
