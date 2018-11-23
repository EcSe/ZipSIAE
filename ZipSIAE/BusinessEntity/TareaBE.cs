using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessEntity
{
    public class TareaBE
    {
        public String IdTarea { get; set; }
        public EntidadDetalleBE TipoTarea { get; set; }
        public String TipoTarea_ValorCadena1 { get { return TipoTarea.ValorCadena1; } }
        public String IdIsoNodo { get; set; }
        public EntidadDetalleBE Contratista { get; set; }
        public DateTime InicioInstalacion { get; set; }
        public DateTime FinInstalacion { get; set; }
        public EntidadDetalleBE Proyecto { get; set; }
        public EntidadDetalleBE TipoNodoA { get; set; }
        public NodoBE NodoIIBBA { get; set; }
        public String NodoIIBBA_IdNodo { get { return NodoIIBBA.IdNodo; } }
        public EntidadDetalleBE TipoNodoB { get; set; }
        public NodoBE NodoB { get; set; }
        public String IdSectorAP { get; set; }
        public Int32 Sector { get; set; }
        public UsuarioBE UsuarioCreacion { get; set; }
        public UsuarioBE UsuarioModificacion { get; set; }
        //PARA EL TIPO DE NODO
        public String Tarea_Tipo_NodoA { get { return TipoNodoA.ValorCadena1; } }
        public TareaBE()
        {
            IdTarea = String.Empty;
            TipoTarea = new EntidadDetalleBE();
            Contratista = new EntidadDetalleBE();
            Proyecto = new EntidadDetalleBE();
            TipoNodoA = new EntidadDetalleBE();
            NodoIIBBA = new NodoBE();
            TipoNodoB = new EntidadDetalleBE();
            NodoB = new NodoBE();
            UsuarioCreacion = new UsuarioBE();
            UsuarioModificacion = new UsuarioBE();
        }
    }
}
