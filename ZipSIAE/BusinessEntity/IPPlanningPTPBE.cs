using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessEntity
{
    public class IPPlanningPTPBE
    {
        public NodoBE NodoA { get; set; }
        //public String Mascara { get; set; }
        public String DefaultGateway { get; set; }
        public String IPConexionLocal { get; set; }
        public UsuarioBE UsuarioCreacion { get; set; }
        public UsuarioBE UsuarioModificacion { get; set; }

        public List<IPPlanningPTPDetalleBE> Detalles { get; set; }
        public IPPlanningPTPBE()
        {
            NodoA = new NodoBE();
            UsuarioCreacion = new UsuarioBE();
            UsuarioModificacion = new UsuarioBE();
            Detalles = new List<IPPlanningPTPDetalleBE>();
        }
    }
}
