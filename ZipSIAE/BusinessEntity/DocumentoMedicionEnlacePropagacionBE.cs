using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessEntity
{
    public class DocumentoMedicionEnlacePropagacionBE
    {
        public DocumentoBE Documento { get; set; }
        //public PMPDetalleBE PMPDetalle { get; set; }
        public NodoBE NodoA { get; set; }
        public String NodoA_IdNodo { get { return NodoA.IdNodo; } }
        public NodoBE NodoIIBBB { get; set; }
        public String NodoIIBBB_IdNodo { get { return NodoIIBBB.IdNodo; } }
        public Double RSSLocal { get; set; }
        public Double RSSRemoto { get; set; }
        public Double TiempoPromedio { get; set; } //cambiado para que acepte decimales
        public Double CapidadSubida { get; set; }
        public Double CapidadBajada { get; set; }
        public UsuarioBE UsuarioCreacion { get; set; }
        public UsuarioBE UsuarioModificacion { get; set; }
        public DocumentoMedicionEnlacePropagacionBE()
        {
            //PMPDetalle = new PMPDetalleBE();
            Documento = new DocumentoBE();
            NodoA = new NodoBE();
            NodoIIBBB = new NodoBE();
            UsuarioCreacion = new UsuarioBE();
            UsuarioModificacion = new UsuarioBE();
        }
    }
}
