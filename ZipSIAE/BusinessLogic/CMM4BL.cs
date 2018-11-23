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
    public class CMM4BL
    {
        public static List<CMM4BE> ListarCMM4(CMM4BE CMM4)
        {
            List<CMM4BE> lstResultados = new List<CMM4BE>();
            DBBaseDatos baseDatos = new DBBaseDatos();
            baseDatos.Configurar();
            baseDatos.Conectar();
            try
            {
                baseDatos.CrearComando("USP_CMM4", CommandType.StoredProcedure);
                baseDatos.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "S", true);
                baseDatos.AsignarParametroCadena("@PCH_ID_NODO", CMM4.Nodo.IdNodo, true);
                
                DbDataReader drDatos = baseDatos.EjecutarConsulta();

                while (drDatos.Read())
                {
                    CMM4BE item = new CMM4BE();
                    item.Nodo.IdNodo = drDatos.GetString(drDatos.GetOrdinal("CH_ID_NODO"));
                    lstResultados.Add(item);
                }

                drDatos.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                baseDatos.Desconectar();
                baseDatos = null;
            }

            return lstResultados;
        }
    }
}
