using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessEntity
{
    public class TareaInstitucioBeneficiariaBE
    {
        public TareaBE Tarea { get; set; }
        public InstitucionBeneficiariaBE InstitucionBeneficiaria { get; set; }
        public String CodigoIIBB { get; set; }

        public TareaInstitucioBeneficiariaBE()
        {
            //Tarea = new TareaBE();
            InstitucionBeneficiaria = new InstitucionBeneficiariaBE();
        }
    }
}
