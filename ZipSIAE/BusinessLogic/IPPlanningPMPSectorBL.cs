using System;
using System.Collections.Generic;
using BusinessEntity;
using DataAccess;
using System.Data;
using System.Data.Common;


namespace BusinessLogic
{
    public class IPPlanningPMPSectorBL
    {
        public static List<IPPlanningPMPSectorBE> ListarIPPlanningPMPSector(IPPlanningPMPSectorBE IPPlanningPMPSector)
        {
            List<IPPlanningPMPSectorBE> lstResultadosBE = new List<IPPlanningPMPSectorBE>();
            DBBaseDatos baseDatosDA = new DBBaseDatos();
            baseDatosDA.Configurar();
            baseDatosDA.Conectar();
            try
            {
                baseDatosDA.CrearComando("USP_IP_PLANNING_PMP_SECTOR", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "S", true);
                if (IPPlanningPMPSector.IPPlanningPMP.Nodo.IdNodo.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PCH_ID_NODO", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PCH_ID_NODO", IPPlanningPMPSector.IPPlanningPMP.Nodo.IdNodo, true);
                baseDatosDA.AsignarParametroEntero("@PIN_SECTOR", IPPlanningPMPSector.SectorNodo, true);

                DbDataReader drDatos = baseDatosDA.EjecutarConsulta();

                while (drDatos.Read())
                {
                    IPPlanningPMPSectorBE item = new IPPlanningPMPSectorBE();

                    item.IPPlanningPMP.Nodo.IdNodo = drDatos.GetString(drDatos.GetOrdinal("CH_ID_NODO"));
                    item.SectorNodo = drDatos.GetInt32(drDatos.GetOrdinal("IN_SECTOR"));
                    item.IPNodo = drDatos.GetString(drDatos.GetOrdinal("VC_IP_NODO"));

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
