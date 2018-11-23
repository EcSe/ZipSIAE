using System;
using System.Collections.Generic;
using System.Text;

namespace BusinessEntity
{
    public class UsuarioBE
    {
        public String IdUsuario { get; set; }
        public EntidadDetalleBE TipoDocumento  { get; set; }
        public String NumeroDocumento { get; set; }
        public String NombreRazon { get; set; }
        public String ApellidoPaterno { get; set; }
        public String ApellidoMaterno { get; set; }
        public String NombreCompleto { get; set; }
        public String Password { get; set; }
        public EntidadDetalleBE Perfil { get; set; }
        public String Email { get; set; }
        public EntidadDetalleBE Contratista { get; set; }
        public UsuarioBE UsuarioCreacion { get; set; }
        public UsuarioBE UsuarioModificacion { get; set; }
        public String Ticket { get; set; }
        public String Metodo { get; set; }

        public UsuarioBE()
        {
            IdUsuario = String.Empty;
            Perfil = new EntidadDetalleBE();
            TipoDocumento = new EntidadDetalleBE();
            Contratista = new EntidadDetalleBE();
            NumeroDocumento = String.Empty;
            NombreRazon = String.Empty;
            ApellidoPaterno = String.Empty;
            ApellidoMaterno = String.Empty;
            NombreCompleto = String.Empty;
            Password = String.Empty;
            Email = String.Empty;
            //UsuarioCreacion = new UsuarioBE();
            Metodo = string.Empty;
        }
    }
}
