using System;
using System.Collections.Generic;
using BusinessEntity;
using DataAccess;
using System.Data;
using System.Data.Common;

namespace BusinessLogic
{
    public class IPPlanningPMPBL
    {
        public static List<IPPlanningPMPBE> ListarIPPlanningPMP(IPPlanningPMPBE IPPlanningPMP)
        {
            List<IPPlanningPMPBE> lstResultadosBE = new List<IPPlanningPMPBE>();
            DBBaseDatos baseDatosDA = new DBBaseDatos();
            baseDatosDA.Configurar();
            baseDatosDA.Conectar();
            try
            {
                baseDatosDA.CrearComando("USP_IP_PLANNING_PMP", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "S", true);
                if (IPPlanningPMP.Nodo.IdNodo.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PCH_ID_NODO", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PCH_ID_NODO", IPPlanningPMP.Nodo.IdNodo, true);

                DbDataReader drDatos = baseDatosDA.EjecutarConsulta();

                while (drDatos.Read())
                {
                    IPPlanningPMPBE item = new IPPlanningPMPBE();

                    item.Nodo.IdNodo = drDatos.GetString(drDatos.GetOrdinal("CH_ID_NODO"));
                    item.DefaultGateway = drDatos.GetString(drDatos.GetOrdinal("VC_DEF_GATEWAY"));

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

        public static void InsertarIPPlanningPMPProceso(IPPlanningPMPBE IPPlanningPMP, DBBaseDatos BaseDatos = null)
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
                baseDatosDA.CrearComando("USP_IP_PLANNING_PMP_PROC", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "I", true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_NODO", IPPlanningPMP.Nodo.IdNodo, true);
                baseDatosDA.AsignarParametroCadena("@PVC_IP_NODO", IPPlanningPMP.IPNodo, true);
                //baseDatosDA.AsignarParametroCadena("@PVC_MASCARA", IPPlanningPMPDetalleBE.IPPlanningPMP.Mascara, true);
                baseDatosDA.AsignarParametroCadena("@PVC_DEF_GATEWAY", IPPlanningPMP.DefaultGateway, true);
                baseDatosDA.AsignarParametroCadena("@PVC_IP_CON_LOCAL", IPPlanningPMP.IPConexionLocal, true);
                baseDatosDA.AsignarParametroCadena("@PVC_PUERTO_NODO", IPPlanningPMP.PuertoNodo, true);
                //baseDatosDA.AsignarParametroCadena("@PVC_PUERTO_IIBB", IPPlanningPMP.PuertoIIBB, true);

                baseDatosDA.AsignarParametroCadena("@PCH_ID_IIBB", IPPlanningPMP.Detalles[0].InstitucionBeneficiaria.IdInstitucionBeneficiaria, true);
                baseDatosDA.AsignarParametroCadena("@PVC_IP_IIBB", IPPlanningPMP.Detalles[0].IPIIBB, true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_USUARIO_CRE", IPPlanningPMP.UsuarioCreacion.IdUsuario, true);

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

        public static void EliminarFisicoIPPlanningPMPProceso(IPPlanningPMPBE IPPlanningPMP, DBBaseDatos BaseDatos = null)
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
                baseDatosDA.CrearComando("USP_IP_PLANNING_PMP_PROC", CommandType.StoredProcedure);
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
