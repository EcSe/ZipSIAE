using BusinessEntity;
using DataAccess;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Text;

namespace BusinessLogic
{
    public class EntidadDetalleBL
    {
        public static List<EntidadDetalleBE> ListarEntidadDetalle(EntidadDetalleBE entidadDetalleBE)
        {
            List<EntidadDetalleBE> lstResultadosBE = new List<EntidadDetalleBE>();
            DBBaseDatos baseDatosDA = new DBBaseDatos();
            baseDatosDA.Configurar();
            baseDatosDA.Conectar();
            try
            {
                baseDatosDA.CrearComando("USP_ENTIDAD_DET", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "S", true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_ENTIDAD_P", entidadDetalleBE.Entidad.IdEntidad, true);
                if (entidadDetalleBE.IdValor.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PVC_ID_VALOR_P", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PVC_ID_VALOR_P", entidadDetalleBE.IdValor, true);
                if (entidadDetalleBE.EntidadDetalleSecundario == null || entidadDetalleBE.EntidadDetalleSecundario.Entidad.IdEntidad.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PVC_ID_ENTIDAD_S", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PVC_ID_ENTIDAD_S", entidadDetalleBE.EntidadDetalleSecundario.Entidad.IdEntidad, true);
                if (entidadDetalleBE.EntidadDetalleSecundario == null || entidadDetalleBE.EntidadDetalleSecundario.IdValor.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PVC_ID_VALOR_S", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PVC_ID_VALOR_S", entidadDetalleBE.EntidadDetalleSecundario.IdValor, true);
                if (entidadDetalleBE.Metodo.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PCH_ID_METODO", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PCH_ID_METODO", entidadDetalleBE.Metodo, true);
                if (entidadDetalleBE.ValorFecha1.Equals(Convert.ToDateTime("01/01/0001 00:00:00.00")))
                    baseDatosDA.AsignarParametroNulo("@PDT_VALOR_FECHA1", true);
                else
                    baseDatosDA.AsignarParametroFecha("@PDT_VALOR_FECHA1", entidadDetalleBE.ValorFecha1, true);

                DbDataReader drDatos = baseDatosDA.EjecutarConsulta();

                while (drDatos.Read())
                {
                    EntidadDetalleBE item = new EntidadDetalleBE();

                    item.Entidad.IdEntidad = drDatos.GetString(drDatos.GetOrdinal("VC_ID_ENTIDAD"));
                    item.IdValor = drDatos.GetString(drDatos.GetOrdinal("VC_ID_VALOR"));
                    if (!drDatos.IsDBNull(drDatos.GetOrdinal("VC_VALOR_CADENA1")))
                        item.ValorCadena1 = drDatos.GetString(drDatos.GetOrdinal("VC_VALOR_CADENA1"));
                    if (!drDatos.IsDBNull(drDatos.GetOrdinal("VC_VALOR_CADENA2")))
                        item.ValorCadena2 = drDatos.GetString(drDatos.GetOrdinal("VC_VALOR_CADENA2"));
                    if (!drDatos.IsDBNull(drDatos.GetOrdinal("VC_VALOR_CADENA3")))
                        item.ValorCadena3 = drDatos.GetString(drDatos.GetOrdinal("VC_VALOR_CADENA3"));
                    if (!drDatos.IsDBNull(drDatos.GetOrdinal("VC_VALOR_CADENA4")))
                        item.ValorCadena4 = drDatos.GetString(drDatos.GetOrdinal("VC_VALOR_CADENA4"));
                    if (!drDatos.IsDBNull(drDatos.GetOrdinal("IN_VALOR_ENTERO1")))
                        item.ValorEntero1 = drDatos.GetInt32(drDatos.GetOrdinal("IN_VALOR_ENTERO1"));
                    if (!drDatos.IsDBNull(drDatos.GetOrdinal("NU_VALOR_NUMERICO1")))
                        item.ValorNumerico1 = (Double)drDatos.GetDecimal(drDatos.GetOrdinal("NU_VALOR_NUMERICO1"));
                    if (!drDatos.IsDBNull(drDatos.GetOrdinal("VB_VALOR_BINARIO1")))
                        item.ValorBinario1 = (Byte[])drDatos.GetValue(drDatos.GetOrdinal("VB_VALOR_BINARIO1"));
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

        public static void InsertarEntidadDetalle(EntidadDetalleBE entidadDetalleBE)
        {
            DBBaseDatos baseDatosDA = new DBBaseDatos();
            baseDatosDA.Configurar();
            baseDatosDA.Conectar();
            try
            {
                baseDatosDA.CrearComando("USP_ENTIDAD_DET", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "I", true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_ENTIDAD_P", entidadDetalleBE.Entidad.IdEntidad, true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_VALOR_P", entidadDetalleBE.IdValor, true);
                if (entidadDetalleBE.ValorCadena1.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PVC_VALOR_CADENA1", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PVC_VALOR_CADENA1", entidadDetalleBE.ValorCadena1, true);
                if (entidadDetalleBE.ValorBinario1 == null || entidadDetalleBE.ValorBinario1.Length.Equals(0))
                    baseDatosDA.AsignarParametroNulo("@PVB_VALOR_BINARIO1", true,ParameterDirection.Input,DbType.Binary);
                else
                    baseDatosDA.AsignarParametroArrayByte("@PVB_VALOR_BINARIO1", entidadDetalleBE.ValorBinario1, true,ParameterDirection.Input,DbType.Binary);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_USUARIO_CRE", entidadDetalleBE.UsuarioCreacion.IdUsuario, true);

                baseDatosDA.EjecutarComando();
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

        }

    }
}
