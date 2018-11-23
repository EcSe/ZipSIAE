using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessEntity
{
    public class TareaNodoBE
    {
        public TareaBE Tarea { get; set; }
        public NodoBE Nodo { get; set; }
        public EntidadDetalleBE TipoNodo { get; set; }
        public String CodigoNodo { get; set; }
        public TareaNodoBE()
        {
            //Tarea = new TareaBE();
            Nodo = new NodoBE();
            TipoNodo = new EntidadDetalleBE();
        }
    }
}
