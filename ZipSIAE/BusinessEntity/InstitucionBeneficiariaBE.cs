using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessEntity
{
    public class InstitucionBeneficiariaBE
    {
        public String IdInstitucionBeneficiaria { get; set; }
        public String Nombre { get; set; }
        public Double Latitud { get; set; }
        public Double Longitud { get; set; }
        public UsuarioBE UsuarioCreacion { get; set; }
        public UsuarioBE UsuarioModificacion { get; set; }

        public InstitucionBeneficiariaBE()
        {
            IdInstitucionBeneficiaria = String.Empty;
            UsuarioCreacion = new UsuarioBE();
            UsuarioModificacion = new UsuarioBE();
        }
    }
}
