using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessEntity
{
    public class PTPBE
    {
        public NodoBE NodoA { get; set; }
        public Double CotaNodoA { get; set; }
        public UsuarioBE UsuarioCreacion { get; set; }
        public UsuarioBE UsuarioModificacion { get; set; }
        public List<PTPDetalleBE> Detalles { get; set; }
        public PTPBE()
        {
            NodoA = new NodoBE();
            Detalles = new List<PTPDetalleBE>();
            UsuarioCreacion = new UsuarioBE();
            UsuarioModificacion = new UsuarioBE();
        }
    }
}
