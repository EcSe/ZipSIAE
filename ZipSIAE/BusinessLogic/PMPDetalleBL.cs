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
    public class PMPDetalleBL
    {
        public static List<PMPDetalleBE> ListarPMPDetalles(PMPDetalleBE PMPDetalle)
        {
            List<PMPDetalleBE> lstResultadosBE = new List<PMPDetalleBE>();
            DBBaseDatos baseDatosDA = new DBBaseDatos();
            baseDatosDA.Configurar();
            baseDatosDA.Conectar();
            try
            {
                baseDatosDA.CrearComando("USP_PMP_DET", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "S", true);

                if (PMPDetalle.PMP.Nodo.IdNodo.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PCH_ID_NODO", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PCH_ID_NODO", PMPDetalle.PMP.Nodo.IdNodo, true);
                if (PMPDetalle.InstitucionBeneficiaria.IdInstitucionBeneficiaria.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PCH_ID_IIBB", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PCH_ID_IIBB", PMPDetalle.InstitucionBeneficiaria.IdInstitucionBeneficiaria, true);
                if (PMPDetalle.SectorIIBB.Equals(0))
                    baseDatosDA.AsignarParametroNulo("@PIN_SECTOR_IIBB", true);
                else
                    baseDatosDA.AsignarParametroEntero("@PIN_SECTOR_IIBB", PMPDetalle.SectorIIBB, true);

                DbDataReader drDatos = baseDatosDA.EjecutarConsulta();

                while (drDatos.Read())
                {
                    PMPDetalleBE item = new PMPDetalleBE();

                    item.PMP.Nodo.IdNodo = drDatos.GetString(drDatos.GetOrdinal("CH_ID_NODO"));
                    item.InstitucionBeneficiaria.IdInstitucionBeneficiaria = drDatos.GetString(drDatos.GetOrdinal("CH_ID_IIBB"));
                    item.SectorIIBB = drDatos.GetInt32(drDatos.GetOrdinal("IN_SECTOR_IIBB"));
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
