using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessEntity
{
    public class KitSIAEBE
    {
        public String SerieKit { get; set; }
        public String CodigoGilat { get; set; }
        public TareaBE Tarea { get; set; }
        public UsuarioBE UsuarioCreacion { get; set; }
        public UsuarioBE UsuarioModificacion { get; set; }
        public KitSIAEBE()
        {
            SerieKit = String.Empty;
            CodigoGilat = String.Empty;
            Tarea = new TareaBE();
            UsuarioCreacion = new UsuarioBE();
            UsuarioModificacion = new UsuarioBE();
        }
    }
}
