using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessEntity
{
    public class PMPDetalleBE
    {
        public PMPBE PMP { get; set; }
        public InstitucionBeneficiariaBE InstitucionBeneficiaria { get; set; }
        public Int32 SectorIIBB { get; set; }
        public Double AzimuthAntenaIIBB { get; set; }
        public Double ElevacionAntenaIIBB { get; set; }
        public Int32 TXTorreIIBB { get; set; }
        public Double EIRPAntenaIIBB { get; set; }
        public Double NivelRXNodo { get; set; }
        public Double NivelRXIIBB { get; set; }
        public Double FadeMarginNodo { get; set; }
        public Double FadeMarginIIBB { get; set; }
        public Double Disponibilidad { get; set; }
        public Double Distancia { get; set; }

        public PMPDetalleBE()
        {
            PMP = new PMPBE();
            InstitucionBeneficiaria = new InstitucionBeneficiariaBE();
        }

    }
}
