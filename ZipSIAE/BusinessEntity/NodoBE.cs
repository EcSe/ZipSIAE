using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessEntity
{
    public class NodoBE
    {
        public String IdNodo { get; set; }
        public DepartamentoBE Region { get; set; }
        public String Nombre { get; set; }
        public CentroPobladoBE Localidad { get; set; }
        //public String Ubigeo { get; set; }
        public Double Latitud { get; set; }
        public Double Longitud { get; set; }
        public Int32 Anillo { get; set; }
        public Int32 AlturaTorre { get; set; }
        public UsuarioBE UsuarioCreacion { get; set; }
        public UsuarioBE UsuarioModificacion { get; set; }

        public NodoBE()
        {
            IdNodo = String.Empty;
            Region = new DepartamentoBE();
            Localidad = new CentroPobladoBE();
            UsuarioCreacion = new UsuarioBE();
            UsuarioModificacion = new UsuarioBE();
        }
    }
}
