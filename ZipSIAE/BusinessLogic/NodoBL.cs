using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BusinessEntity;
using DataAccess;
using System.Data;
using System.Data.Common;

namespace BusinessLogic
{
    public class NodoBL
    {
        public static void InsertarNodoProceso(NodoBE nodoBE, DBBaseDatos BaseDatos = null)
        {
            DBBaseDatos baseDatosDA = new DBBaseDatos();
            if (BaseDatos == null)
            {
                baseDatosDA.Configurar();
                baseDatosDA.Conectar();
            }
            else
            {
                baseDatosDA = BaseDatos;
            }

            try
            {
                baseDatosDA.CrearComando("USP_NODO_PROC", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "I", true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_NODO", nodoBE.IdNodo, true);
                baseDatosDA.AsignarParametroCadena("@PVC_NOMBRE_REGION", nodoBE.Region.Nombre, true);
                baseDatosDA.AsignarParametroCadena("@PVC_NOMBRE", nodoBE.Nombre, true);
                baseDatosDA.AsignarParametroCadena("@PVC_NOMBRE_DEPARTAMENTO", nodoBE.Localidad.Distrito.Provincia.Departamento.Nombre, true);
                baseDatosDA.AsignarParametroCadena("@PVC_NOMBRE_PROVINCIA", nodoBE.Localidad.Distrito.Provincia.Nombre, true);
                baseDatosDA.AsignarParametroCadena("@PVC_NOMBRE_DISTRITO", nodoBE.Localidad.Distrito.Nombre, true);
                baseDatosDA.AsignarParametroCadena("@PVC_NOMBRE_CCPP", nodoBE.Localidad.Nombre, true);
                baseDatosDA.AsignarParametroDouble("@PNU_LATITUD", nodoBE.Latitud, true);
                baseDatosDA.AsignarParametroDouble("@PNU_LONGITUD", nodoBE.Longitud, true);
                baseDatosDA.AsignarParametroEntero("@PIN_ANILLO", nodoBE.Anillo, true);
                baseDatosDA.AsignarParametroEntero("@PIN_ALT_TORRE", nodoBE.AlturaTorre, true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_USUARIO_CRE", nodoBE.UsuarioCreacion.IdUsuario, true);
                baseDatosDA.EjecutarComando();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (BaseDatos == null)
                {
                    baseDatosDA.Desconectar();
                    baseDatosDA = null;
                }
            }

            
        }

        public static void EliminarFisicoNodoProceso(NodoBE nodoBE, DBBaseDatos BaseDatos = null)
        {
            DBBaseDatos baseDatosDA = new DBBaseDatos();
            if (BaseDatos == null)
            {
                baseDatosDA.Configurar();
                baseDatosDA.Conectar();
            }
            else
            {
                baseDatosDA = BaseDatos;
            }

            try
            {
                baseDatosDA.CrearComando("USP_NODO_PROC", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "F", true);
                baseDatosDA.EjecutarComando();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (BaseDatos == null)
                {
                    baseDatosDA.Desconectar();
                    baseDatosDA = null;
                }
            }

        }

        public static List<NodoBE> ListarNodos(NodoBE nodoBE)
        {
            List<NodoBE> lstResultadosBE = new List<NodoBE>();
            DBBaseDatos baseDatosDA = new DBBaseDatos();
            baseDatosDA.Configurar();
            baseDatosDA.Conectar();
            try
            {
                baseDatosDA.CrearComando("USP_NODO", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "S", true);

                if (!nodoBE.IdNodo.Equals(""))
                    baseDatosDA.AsignarParametroCadena("@PCH_ID_NODO", nodoBE.IdNodo, true);
                else
                    baseDatosDA.AsignarParametroNulo("@PCH_ID_NODO", true);

                DbDataReader drDatos = baseDatosDA.EjecutarConsulta();

                while (drDatos.Read())
                {
                    NodoBE item = new NodoBE();

                    item.IdNodo = drDatos.GetString(drDatos.GetOrdinal("CH_ID_NODO"));
                    item.Nombre = drDatos.GetString(drDatos.GetOrdinal("VC_NOMBRE"));
                    item.Region.Nombre = drDatos.GetString(drDatos.GetOrdinal("VC_NOMBRE_REGION"));
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
