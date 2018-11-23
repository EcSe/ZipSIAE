using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessEntity
{
    public class OpcionBE
    {
        public String IdOpcion { get; set; }
        public String Nombre { get; set; }
        public String IdOpcionPadre { get; set; }
        public String URL { get; set; }
        public String Icono { get; set; }
        public String Orden { get; set; }
    }
}
