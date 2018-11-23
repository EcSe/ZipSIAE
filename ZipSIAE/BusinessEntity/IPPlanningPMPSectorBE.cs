using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessEntity
{
    public class IPPlanningPMPSectorBE
    {
        public IPPlanningPMPBE IPPlanningPMP { get; set; }
        public Int32 SectorNodo { get; set; }
        public String IPNodo { get; set; }

        public IPPlanningPMPSectorBE()
        {
            IPPlanningPMP = new IPPlanningPMPBE();
            IPNodo = String.Empty;
        }
    }
}
