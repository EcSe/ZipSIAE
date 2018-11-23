using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessEntity
{
    public class DocumentoEquipamientoBE
    {
        public DocumentoBE Documento { get; set; }
        public EntidadDetalleBE Equipamiento { get; set; }
        public String SerieEquipamiento { get; set; }
        public Int32 Item { get; set; }
        public String IdEmpresa { get; set; }
        public KitSIAEBE KitSIAE { get; set; }
        public String Equipamiento_IdValor { get { return Equipamiento.IdValor; } }
        public UsuarioBE UsuarioCreacion { get; set; }
        public UsuarioBE UsuarioModificacion { get; set; }

        public DocumentoEquipamientoBE()
        {
            Documento = new DocumentoBE();
            Equipamiento = new EntidadDetalleBE();
            SerieEquipamiento = String.Empty;
            KitSIAE = new KitSIAEBE();
            UsuarioCreacion = new UsuarioBE();
            UsuarioModificacion = new UsuarioBE();
        }

    }
}
