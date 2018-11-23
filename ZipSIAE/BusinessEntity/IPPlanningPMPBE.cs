using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessEntity
{
    public class IPPlanningPMPBE
    {
        public NodoBE Nodo { get; set; }
        public String IPNodo { get; set; }
        //public String Mascara { get; set; }
        public String DefaultGateway { get; set; }
        public String IPConexionLocal { get; set; }
        public String PuertoNodo { get; set; }
        //public String PuertoIIBB { get; set; }
        public UsuarioBE UsuarioCreacion { get; set; }
        public UsuarioBE UsuarioModificacion { get; set; }
        public List<IPPlanningPMPDetalleBE> Detalles { get; set; }
        public IPPlanningPMPBE()
        {
            Nodo = new NodoBE();
            Detalles = new List<IPPlanningPMPDetalleBE>();
            UsuarioCreacion = new UsuarioBE();
            UsuarioModificacion = new UsuarioBE();
        }
    }
}
