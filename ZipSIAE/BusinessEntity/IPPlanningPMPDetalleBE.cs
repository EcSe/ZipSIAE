using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessEntity
{
    public class IPPlanningPMPDetalleBE
    {
        public IPPlanningPMPBE IPPlanningPMP { get; set; }
        public InstitucionBeneficiariaBE InstitucionBeneficiaria { get; set; }
        public String IPIIBB { get; set; }

        public IPPlanningPMPDetalleBE()
        {
            IPPlanningPMP = new IPPlanningPMPBE();
            InstitucionBeneficiaria = new InstitucionBeneficiariaBE();
        }
    }
}
