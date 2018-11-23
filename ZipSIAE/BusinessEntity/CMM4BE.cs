using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessEntity
{
    public class CMM4BE
    {
        public NodoBE Nodo { get; set; }

        public CMM4BE()
        {
            Nodo = new NodoBE();
        }
    }
}
