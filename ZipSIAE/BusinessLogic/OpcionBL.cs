using BusinessEntity;
using DataAccess;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessLogic
{
    public class OpcionBL
    {
        public static List<OpcionBE> ListarOpciones(UsuarioBE usuarioBE, AplicacionBE aplicacionBE)
        {
            List<OpcionBE> lstResultadosBE = new List<OpcionBE>();
            DBBaseDatos baseDatosDA = new DBBaseDatos();
            baseDatosDA.Configurar();
            baseDatosDA.Conectar();

            try
            {
                //baseDatosDA.CrearComando("USP_LISTAR_OPCIONES", CommandType.StoredProcedure);
                baseDatosDA.CrearComando("USP_OPCION", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "S", true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_USUARIO", usuarioBE.IdUsuario, true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_PERFIL", usuarioBE.Perfil.IdValor, true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_APLICACION", aplicacionBE.IdAplicacion, true);

                DbDataReader drDatos = baseDatosDA.EjecutarConsulta();

                while (drDatos.Read())
                {
                    OpcionBE item = new OpcionBE();

                    item.IdOpcion = drDatos.GetString(drDatos.GetOrdinal("VC_ID_OPCION"));
                    item.Nombre = drDatos.GetString(drDatos.GetOrdinal("VC_NOMBRE"));
                    if (!drDatos.IsDBNull(drDatos.GetOrdinal("VC_ID_OPCION_PADRE")))
                        item.IdOpcionPadre = drDatos.GetString(drDatos.GetOrdinal("VC_ID_OPCION_PADRE"));
                    if (!drDatos.IsDBNull(drDatos.GetOrdinal("VC_URL")))
                        item.URL = drDatos.GetString(drDatos.GetOrdinal("VC_URL"));
                    item.Icono = drDatos.GetString(drDatos.GetOrdinal("VC_ICONO"));
                    item.Orden = drDatos.GetString(drDatos.GetOrdinal("VC_ORDEN"));

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
