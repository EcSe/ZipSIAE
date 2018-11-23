using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessEntity
{
    public class CentroPobladoBE
    {
        public DistritoBE Distrito { get; set; }
        public String IdCentroPoblado { get; set; }
        public String Nombre { get; set; }

        public CentroPobladoBE()
        {
            Distrito = new DistritoBE();
        }
    }
}
