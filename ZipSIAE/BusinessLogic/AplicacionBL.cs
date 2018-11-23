using System;
using System.Collections.Generic;
using System.Text;
using BusinessEntity;
using DataAccess;
using System.Data;
using System.Data.Common;

namespace BusinessLogic
{
    public class AplicacionBL
    {
        public static List<AplicacionBE> ListarAplicaciones(UsuarioBE usuarioBE)
        {
            List<AplicacionBE> lstResultadosBE = new List<AplicacionBE>();
            DBBaseDatos baseDatosDA = new DBBaseDatos();
            baseDatosDA.Configurar();
            baseDatosDA.Conectar();
            try
            {
                //baseDatosDA.CrearComando("USP_LISTAR_APLICACIONES", CommandType.StoredProcedure);
                baseDatosDA.CrearComando("USP_APLICACION", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "S", true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_USUARIO", usuarioBE.IdUsuario, true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_PERFIL", usuarioBE.Perfil.IdValor, true);

                DbDataReader drDatos = baseDatosDA.EjecutarConsulta();

                while (drDatos.Read())
                {
                    AplicacionBE item = new AplicacionBE();

                    item.IdAplicacion = drDatos.GetString(drDatos.GetOrdinal("VC_ID_APLICACION"));
                    item.Nombre = drDatos.GetString(drDatos.GetOrdinal("VC_NOMBRE"));
                    if (!drDatos.IsDBNull(drDatos.GetOrdinal("VC_DESCRIPCION")))
                        item.Descripcion = drDatos.GetString(drDatos.GetOrdinal("VC_DESCRIPCION"));
                    item.URLDefault = drDatos.GetString(drDatos.GetOrdinal("VC_URL_DEFAULT"));
                    item.Icono = drDatos.GetString(drDatos.GetOrdinal("VC_ICONO"));
                    item.EstiloIcono = drDatos.GetString(drDatos.GetOrdinal("VC_ESTILO_ICONO"));
                    item.EstiloTitulo = drDatos.GetString(drDatos.GetOrdinal("VC_ESTILO_TITULO"));
                    item.EstiloBoton = drDatos.GetString(drDatos.GetOrdinal("VC_ESTILO_BOTON"));

                    lstResultadosBE.Add(item);
                }
                drDatos.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                baseDatosDA.Desconectar();
                baseDatosDA = null;
            }
            
            return lstResultadosBE;
        }
    }
}
