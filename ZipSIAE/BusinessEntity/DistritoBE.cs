using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessEntity
{
    public class DistritoBE
    {
        public ProvinciaBE Provincia { get; set; }
        public String IdDistrito { get; set; }
        public String Nombre { get; set; }

        public DistritoBE()
        {
            Provincia = new ProvinciaBE();
        }
    }
}
