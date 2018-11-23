using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessEntity
{
    public class IPPlanningPTPDetalleBE
    {
        public IPPlanningPTPBE IPPlanningPTP { get; set; }
        public NodoBE NodoB { get; set; }
        public String IPNodoA { get; set; }
        public String IPNodoB { get; set; }
        public String PuertoNodoA { get; set; }
        public String PuertoNodoB { get; set; }
        public Int32 CodigoColor { get; set; }
        public NodoBE NodoMaestro { get; set; }
        public EntidadDetalleBE Sincronismo { get; set; }

        public IPPlanningPTPDetalleBE()
        {
            IPPlanningPTP = new IPPlanningPTPBE();
            NodoB = new NodoBE();
            NodoMaestro = new NodoBE();
            Sincronismo = new EntidadDetalleBE();
        }
    }
}
