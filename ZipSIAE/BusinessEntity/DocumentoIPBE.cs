using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessEntity
{
    public class DocumentoIPBE
    {
        public DocumentoBE Documento { get; set; }
        public String IPSystem { get; set; }
        public String RangoGestionSeguridadEnergia { get; set; }
        public String Gateway { get; set; }
        public String Mascara { get; set; }
        public String IPReservada { get; set; }
        public List<DocumentoIPEquipamientoBE> Equipamientos { get; set; }
        public UsuarioBE UsuarioCreacion { get; set; }
        public UsuarioBE UsuarioModificacion { get; set; }
        public DocumentoIPBE()
        {
            Documento = new DocumentoBE();
            Equipamientos = new List<DocumentoIPEquipamientoBE>();
            UsuarioCreacion = new UsuarioBE();
            UsuarioModificacion = new UsuarioBE();
        }

    }
}
