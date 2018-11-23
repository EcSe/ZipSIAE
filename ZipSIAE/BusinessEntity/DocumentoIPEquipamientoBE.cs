using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessEntity
{
    public class DocumentoIPEquipamientoBE
    {
        public DocumentoIPBE DocumentoIP { get; set; }
        public EntidadDetalleBE Equipamiento { get; set; }
        public String IPEquipamiento { get; set; }
        public UsuarioBE UsuarioCreacion { get; set; }
        public UsuarioBE UsuarioModificacion { get; set; }
        public DocumentoIPEquipamientoBE()
        {
            Equipamiento = new EntidadDetalleBE();
            UsuarioCreacion = new UsuarioBE();
            UsuarioModificacion = new UsuarioBE();
        }
    }
}
