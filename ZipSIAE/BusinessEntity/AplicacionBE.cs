using System;
using System.Collections.Generic;
using System.Text;

namespace BusinessEntity
{
    public class AplicacionBE
    {
        public String IdAplicacion { get; set; }
        public String Nombre { get; set; }
        public String Descripcion { get; set; }
        public String URLDefault { get; set; }
        public String Icono { get; set; }
        public String EstiloIcono { get; set; }
        public String EstiloTitulo { get; set; }
        public String EstiloBoton { get; set; }

        public List<OpcionBE> Opciones { get; set; }

        public AplicacionBE()
        {
            Opciones = new List<OpcionBE>();
        }
    }
}
