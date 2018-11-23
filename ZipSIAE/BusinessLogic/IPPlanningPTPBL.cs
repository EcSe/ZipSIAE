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
    public class IPPlanningPTPBL
    {
        public static void InsertarIPPlanningPTPProceso(IPPlanningPTPBE IPPlanningPTP, DBBaseDatos BaseDatos = null)
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
                baseDatosDA.CrearComando("USP_IP_PLANNING_PTP_PROC", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "I", true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_NODO_A", IPPlanningPTP.NodoA.IdNodo, true);
                baseDatosDA.AsignarParametroCadena("@PVC_DEF_GATEWAY", IPPlanningPTP.DefaultGateway, true);
                baseDatosDA.AsignarParametroCadena("@PVC_IP_CON_LOCAL", IPPlanningPTP.IPConexionLocal, true);

                //if (IPPlanningPTP.Detalles[0].NodoB.IdNodo == null || IPPlanningPTP.Detalles[0].NodoB.IdNodo.Equals(""))
                if (IPPlanningPTP.Detalles[0].NodoB.IdNodo.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PCH_ID_NODO_B", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PCH_ID_NODO_B", IPPlanningPTP.Detalles[0].NodoB.IdNodo, true);
                baseDatosDA.AsignarParametroCadena("@PVC_IP_NODO_A", IPPlanningPTP.Detalles[0].IPNodoA, true);
                if (IPPlanningPTP.Detalles[0].IPNodoB == null || IPPlanningPTP.Detalles[0].IPNodoB.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PVC_IP_NODO_B", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PVC_IP_NODO_B", IPPlanningPTP.Detalles[0].IPNodoB, true);
                baseDatosDA.AsignarParametroCadena("@PVC_PUERTO_NODO_A", IPPlanningPTP.Detalles[0].PuertoNodoA, true);
                if (IPPlanningPTP.Detalles[0].PuertoNodoB == null || IPPlanningPTP.Detalles[0].PuertoNodoB.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PVC_PUERTO_NODO_B", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PVC_PUERTO_NODO_B", IPPlanningPTP.Detalles[0].PuertoNodoB, true);
                if (IPPlanningPTP.Detalles[0].CodigoColor.Equals(0))
                    baseDatosDA.AsignarParametroNulo("@PIN_COD_COLOR", true);
                else
                    baseDatosDA.AsignarParametroEntero("@PIN_COD_COLOR", IPPlanningPTP.Detalles[0].CodigoColor, true);
                //if (IPPlanningPTP.Detalles[0].NodoMaestro.IdNodo==null || IPPlanningPTP.Detalles[0].NodoMaestro.IdNodo.Equals(""))
                if (IPPlanningPTP.Detalles[0].NodoMaestro.IdNodo.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PCH_ID_NODO_MAESTRO", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PCH_ID_NODO_MAESTRO", IPPlanningPTP.Detalles[0].NodoMaestro.IdNodo, true);
                if (IPPlanningPTP.Detalles[0].Sincronismo.ValorCadena1 == null || IPPlanningPTP.Detalles[0].Sincronismo.ValorCadena1.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PVC_NOMBRE_SINCRONISMO", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PVC_NOMBRE_SINCRONISMO", IPPlanningPTP.Detalles[0].Sincronismo.ValorCadena1, true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_USUARIO_CRE", IPPlanningPTP.UsuarioCreacion.IdUsuario, true);
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

        public static void EliminarFisicoIPPlanningPTPProceso(IPPlanningPTPBE IPPlanningPTP, DBBaseDatos BaseDatos = null)
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
                baseDatosDA.CrearComando("USP_IP_PLANNING_PTP_PROC", CommandType.StoredProcedure);
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
