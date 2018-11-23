using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessEntity
{
    
    public class DocumentoMaterialBE
    {
        public DocumentoBE Documento { get; set; }
        public EntidadDetalleBE Material { get; set; }
        public Double Cantidad { get; set; }
        public String Material_IdValor { get { return Material.IdValor; } }
        public UsuarioBE UsuarioCreacion { get; set; }
        public UsuarioBE UsuarioModificacion { get; set; }

        public DocumentoMaterialBE()
        {
            Documento = new DocumentoBE();
            Material = new EntidadDetalleBE();
            UsuarioCreacion = new UsuarioBE();
            UsuarioModificacion = new UsuarioBE();
        }
    }
}
