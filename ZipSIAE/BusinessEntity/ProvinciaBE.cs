using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessEntity
{
    public class ProvinciaBE
    {
        public DepartamentoBE Departamento { get; set; }
        public String IdProvincia { get; set; }
        public string Nombre { get; set; }

        public ProvinciaBE()
        {
            Departamento = new DepartamentoBE();
        }
    }
}
