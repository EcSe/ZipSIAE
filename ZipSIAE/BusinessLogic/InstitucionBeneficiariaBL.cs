using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BusinessEntity;
using DataAccess;
using System.Data;

namespace BusinessLogic
{
    public class InstitucionBeneficiariaBL
    {
        public static void InsertarInstitucionBeneficiariaProceso(InstitucionBeneficiariaBE InstitucionBeneficiaria, DBBaseDatos BaseDatos = null)
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
                baseDatosDA.CrearComando("USP_INSTITUCION_BENEFICIARIA_PROC", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "I", true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_IIBB", InstitucionBeneficiaria.IdInstitucionBeneficiaria, true);
                baseDatosDA.AsignarParametroCadena("@PVC_NOMBRE_IIBB", InstitucionBeneficiaria.Nombre, true);
                baseDatosDA.AsignarParametroDouble("@PNU_LATITUD_IIBB", InstitucionBeneficiaria.Latitud, true);
                baseDatosDA.AsignarParametroDouble("@PNU_LONGITUD_IIBB", InstitucionBeneficiaria.Longitud, true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_USUARIO_CRE", InstitucionBeneficiaria.UsuarioCreacion.IdUsuario, true);
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


        public static void EliminarFisicoInstitucionBeneficiariaProceso(InstitucionBeneficiariaBE InstitucionBeneficiaria, DBBaseDatos BaseDatos = null)
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
                baseDatosDA.CrearComando("USP_INSTITUCION_BENEFICIARIA_PROC", CommandType.StoredProcedure);
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
