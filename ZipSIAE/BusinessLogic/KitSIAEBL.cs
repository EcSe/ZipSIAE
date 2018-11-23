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
    public class KitSIAEBL
    {
        public static void InsertarKitSIAEProceso(KitSIAEBE KitSIAE, DBBaseDatos BaseDatos = null)
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
                baseDatosDA.CrearComando("USP_KIT_SIAE_PROC", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "I", true);
                baseDatosDA.AsignarParametroCadena("@PVC_SERIE_KIT", KitSIAE.SerieKit, true);
                baseDatosDA.AsignarParametroCadena("@PVC_CODIGO_GILAT", KitSIAE.CodigoGilat, true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_TAREA", KitSIAE.Tarea.IdTarea, true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_NODO_IIBB", KitSIAE.Tarea.NodoIIBBA.IdNodo, true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_USUARIO_CRE", KitSIAE.UsuarioCreacion.IdUsuario, true);

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

        public static void EliminarFisicoKitSIAEProceso(KitSIAEBE KitSIAE, DBBaseDatos BaseDatos = null)
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
                baseDatosDA.CrearComando("USP_KIT_SIAE_PROC", CommandType.StoredProcedure);
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
    }
}
