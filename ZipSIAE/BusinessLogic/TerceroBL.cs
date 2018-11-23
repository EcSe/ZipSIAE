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
    public class TerceroBL
    {
        public static List<TerceroBE> ListarTerceros(TerceroBE tercerobE)
        {
            List<TerceroBE> lstResultadosBE = new List<TerceroBE>();
            DBBaseDatos baseDatosDA = new DBBaseDatos();
            baseDatosDA.Configurar();
            baseDatosDA.Conectar();
            try
            {

                baseDatosDA.CrearComando("USP_TERCERO", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "S", true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_ACT", tercerobE.Actividad.IdValor, true);
                baseDatosDA.AsignarParametroCadena("@PVC_NOM_COMP", tercerobE.NombreCompleto, true);

                DbDataReader drDatos = baseDatosDA.EjecutarConsulta();

                while (drDatos.Read())
                {
                    TerceroBE item = new TerceroBE();

                    item.TipoDocumento.IdValor = drDatos.GetString(drDatos.GetOrdinal("CH_ID_TIP_DOC"));
                    item.NumeroDocumento = drDatos.GetString(drDatos.GetOrdinal("VC_NUM_DOC"));
                    item.NombreRazon = drDatos.GetString(drDatos.GetOrdinal("VC_NOMBRE_RAZON"));
                    if (!drDatos.IsDBNull(drDatos.GetOrdinal("VC_APE_PAT")))
                        item.ApellidoPaterno = drDatos.GetString(drDatos.GetOrdinal("VC_APE_PAT"));
                    if (!drDatos.IsDBNull(drDatos.GetOrdinal("VC_APE_MAT")))
                        item.ApellidoMaterno = drDatos.GetString(drDatos.GetOrdinal("VC_APE_MAT"));
                    item.NombreCompleto = drDatos.GetString(drDatos.GetOrdinal("VC_NOM_COMP"));
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
