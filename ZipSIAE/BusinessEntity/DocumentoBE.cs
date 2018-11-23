using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessEntity
{
    public class DocumentoBE
    {
        public EntidadDetalleBE Documento { get; set; }
        public TareaBE Tarea { get; set; }
        //public String Tarea_IdTarea { get { return Tarea.IdTarea; } }
        public String Documento_IdValor { get { return Documento.IdValor; } }
        public String Documento_ValorCadena1 { get { return Documento.ValorCadena1; } }
        public String Documento_ValorCadena2 { get { return Documento.ValorCadena2; } }
       
        public Double PorcentajeAvance { get; set; }
        public Double PorcentajeAprobado { get; set; }
        public List<DocumentoDetalleBE> Detalles { get; set; }
        public List<DocumentoEquipamientoBE> Equipamientos { get; set; }
        public List<DocumentoMaterialBE> Materiales { get; set; }
        public List<DocumentoMedicionEnlacePropagacionBE> MedicionesEnlacePropagacion { get; set; }
        public DocumentoBE()
        {
            Documento = new EntidadDetalleBE();
            Tarea = new TareaBE();
            Detalles = new List<DocumentoDetalleBE>();
            Equipamientos = new List<DocumentoEquipamientoBE>();
            Materiales = new List<DocumentoMaterialBE>();
            MedicionesEnlacePropagacion = new List<DocumentoMedicionEnlacePropagacionBE>();
        }

        public DocumentoBE Clone()
        {
            return (DocumentoBE)this.MemberwiseClone();
        }

    }
}
