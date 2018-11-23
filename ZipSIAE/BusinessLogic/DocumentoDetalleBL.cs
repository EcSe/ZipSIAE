using System;
using System.Collections.Generic;
using BusinessEntity;
using DataAccess;
using System.Data;
using System.Data.Common;


namespace BusinessLogic
{
    public class DocumentoDetalleBL
    {
        public static List<DocumentoDetalleBE> ListarDocumentoDetalle(DocumentoDetalleBE DocumentoDetalle)
        {
            List<DocumentoDetalleBE> lstResultado = new List<DocumentoDetalleBE>();
            DBBaseDatos baseDatosDA = new DBBaseDatos();
            baseDatosDA.Configurar();
            baseDatosDA.Conectar();
            try
            {
                baseDatosDA.CrearComando("USP_DOCUMENTO_DET", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "S", true);
                if (DocumentoDetalle.Documento.Documento.IdValor.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PCH_ID_DOCUMENTO", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PCH_ID_DOCUMENTO", DocumentoDetalle.Documento.Documento.IdValor, true);
                if (DocumentoDetalle.Documento.Tarea.IdTarea.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PCH_ID_TAREA", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PCH_ID_TAREA", DocumentoDetalle.Documento.Tarea.IdTarea, true);

                DbDataReader drDatos = baseDatosDA.EjecutarConsulta();

                while (drDatos.Read())
                {
                    DocumentoDetalleBE item = new DocumentoDetalleBE();
                    item.Documento = new DocumentoBE();

                    item.Documento.Documento.IdValor = drDatos.GetString(drDatos.GetOrdinal("CH_ID_DOCUMENTO"));
                    item.Documento.Tarea.IdTarea = drDatos.GetString(drDatos.GetOrdinal("CH_ID_TAREA"));
                    item.Campo.IdValor = drDatos.GetString(drDatos.GetOrdinal("CH_ID_CAMPO"));

                    if (!drDatos.IsDBNull(drDatos.GetOrdinal("VC_VALOR_CAMPO")))
                        item.IdValor = drDatos.GetString(drDatos.GetOrdinal("VC_VALOR_CAMPO"));
                    if (!drDatos.IsDBNull(drDatos.GetOrdinal("VC_VALOR_CADENA")))
                        item.ValorCadena = drDatos.GetString(drDatos.GetOrdinal("VC_VALOR_CADENA"));
                    if (!drDatos.IsDBNull(drDatos.GetOrdinal("IN_VALOR_ENTERO")))
                        item.ValorEntero = drDatos.GetInt32(drDatos.GetOrdinal("IN_VALOR_ENTERO"));
                    if (!drDatos.IsDBNull(drDatos.GetOrdinal("NU_VALOR_NUMERICO")))
                        item.ValorNumerico = (Double)drDatos.GetDecimal(drDatos.GetOrdinal("NU_VALOR_NUMERICO"));
                    if (!drDatos.IsDBNull(drDatos.GetOrdinal("DT_VALOR_FECHA")))
                        item.ValorFecha = drDatos.GetDateTime(drDatos.GetOrdinal("DT_VALOR_FECHA"));
                    if (!drDatos.IsDBNull(drDatos.GetOrdinal("BL_VALOR_BOOLEANO")))
                        item.ValorBoolean = drDatos.GetBoolean(drDatos.GetOrdinal("BL_VALOR_BOOLEANO"));
                    if (!drDatos.IsDBNull(drDatos.GetOrdinal("VB_VALOR_BINARIO")))
                        item.ValorBinario = (Byte[])drDatos.GetValue(drDatos.GetOrdinal("VB_VALOR_BINARIO"));
                    if (!drDatos.IsDBNull(drDatos.GetOrdinal("VC_EXTENSION_ARCHIVO")))
                        item.ExtensionArchivo = drDatos.GetString(drDatos.GetOrdinal("VC_EXTENSION_ARCHIVO"));
                    item.Aprobado = drDatos.GetBoolean(drDatos.GetOrdinal("BL_APROBADO"));
                    if (!drDatos.IsDBNull(drDatos.GetOrdinal("VC_COMENTARIO")))
                        item.Comentario = drDatos.GetString(drDatos.GetOrdinal("VC_COMENTARIO"));
                    lstResultado.Add(item);
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

            return lstResultado;
        }

        public static void InsertarDocumentoDetalle(DocumentoDetalleBE DocumentoDetalle,DBBaseDatos BaseDatos=null)
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
                baseDatosDA.CrearComando("USP_DOCUMENTO_DET", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "I", true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_DOCUMENTO", DocumentoDetalle.Documento.Documento.IdValor, true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_TAREA", DocumentoDetalle.Documento.Tarea.IdTarea, true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_CAMPO", DocumentoDetalle.Campo.IdValor, true);
                if (DocumentoDetalle.IdValor.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PVC_VALOR_CAMPO", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PVC_VALOR_CAMPO", DocumentoDetalle.IdValor, true);
                if (DocumentoDetalle.ValorCadena.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PVC_VALOR_CADENA", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PVC_VALOR_CADENA", DocumentoDetalle.ValorCadena, true);
                if (DocumentoDetalle.ValorEntero.Equals(null))
                    baseDatosDA.AsignarParametroNulo("@PIN_VALOR_ENTERO", true);
                else
                    baseDatosDA.AsignarParametroEntero("@PIN_VALOR_ENTERO", (Int32)DocumentoDetalle.ValorEntero, true);
                if (DocumentoDetalle.ValorNumerico.Equals(null))
                    baseDatosDA.AsignarParametroNulo("@PNU_VALOR_NUMERICO", true);
                else
                    baseDatosDA.AsignarParametroDouble("@PNU_VALOR_NUMERICO", (Double)DocumentoDetalle.ValorNumerico, true);
                if (DocumentoDetalle.ValorFecha.Equals(Convert.ToDateTime("01/01/0001 00:00:00.00")))
                    baseDatosDA.AsignarParametroNulo("@PDT_VALOR_FECHA", true);
                else
                    baseDatosDA.AsignarParametroFecha("@PDT_VALOR_FECHA", DocumentoDetalle.ValorFecha, true);
                if (DocumentoDetalle.ValorBoolean.Equals(null))
                    baseDatosDA.AsignarParametroNulo("@PBL_VALOR_BOOLEANO", true);
                else
                    baseDatosDA.AsignarParametroBoolean("@PBL_VALOR_BOOLEANO", (Boolean)DocumentoDetalle.ValorBoolean, true);
                if (DocumentoDetalle.ValorBinario == null || DocumentoDetalle.ValorBinario.Length.Equals(0))
                    baseDatosDA.AsignarParametroNulo("@PVB_VALOR_BINARIO", true,ParameterDirection.Input,DbType.Binary);
                else
                    baseDatosDA.AsignarParametroArrayByte("@PVB_VALOR_BINARIO", DocumentoDetalle.ValorBinario, true, ParameterDirection.Input, DbType.Binary);
                if (DocumentoDetalle.ExtensionArchivo.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PVC_EXTENSION_ARCHIVO", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PVC_EXTENSION_ARCHIVO", DocumentoDetalle.ExtensionArchivo, true);
                baseDatosDA.AsignarParametroBoolean("@PBL_APROBADO", DocumentoDetalle.Aprobado, true);
                baseDatosDA.AsignarParametroCadena("@PVC_COMENTARIO", DocumentoDetalle.Comentario, true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_USUARIO_CRE", DocumentoDetalle.UsuarioCreacion.IdUsuario, true);

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

        public static void InsertarDocumentoDetalleProceso(DocumentoDetalleBE DocumentoDetalle, DBBaseDatos BaseDatos = null)
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
                baseDatosDA.CrearComando("USP_DOCUMENTO_DET_PROC", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "I", true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_DOCUMENTO", DocumentoDetalle.Documento.Documento.IdValor, true);
                //baseDatosDA.AsignarParametroCadena("@PCH_ID_TAREA", DocumentoDetalle.Documento.Tarea.IdTarea, true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_NODO_IIBB_A", DocumentoDetalle.Documento.Tarea.NodoIIBBA.IdNodo, true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_CAMPO", DocumentoDetalle.Campo.IdValor, true);
                if (DocumentoDetalle.IdValor.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PVC_VALOR_CAMPO", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PVC_VALOR_CAMPO", DocumentoDetalle.IdValor, true);
                if (DocumentoDetalle.ValorCadena.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PVC_VALOR_CADENA", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PVC_VALOR_CADENA", DocumentoDetalle.ValorCadena, true);
                if (DocumentoDetalle.ValorEntero.Equals(null))
                    baseDatosDA.AsignarParametroNulo("@PIN_VALOR_ENTERO", true);
                else
                    baseDatosDA.AsignarParametroEntero("@PIN_VALOR_ENTERO", (Int32)DocumentoDetalle.ValorEntero, true);
                if (DocumentoDetalle.ValorNumerico.Equals(null))
                    baseDatosDA.AsignarParametroNulo("@PNU_VALOR_NUMERICO", true);
                else
                    baseDatosDA.AsignarParametroDouble("@PNU_VALOR_NUMERICO", (Double)DocumentoDetalle.ValorNumerico, true);
                if (DocumentoDetalle.ValorFecha.Equals(Convert.ToDateTime("01/01/0001 00:00:00.00")))
                    baseDatosDA.AsignarParametroNulo("@PDT_VALOR_FECHA", true);
                else
                    baseDatosDA.AsignarParametroFecha("@PDT_VALOR_FECHA", DocumentoDetalle.ValorFecha, true);
                if (DocumentoDetalle.ValorBoolean.Equals(null))
                    baseDatosDA.AsignarParametroNulo("@PBL_VALOR_BOOLEANO", true);
                else
                    baseDatosDA.AsignarParametroBoolean("@PBL_VALOR_BOOLEANO", (Boolean)DocumentoDetalle.ValorBoolean, true);
                //if (DocumentoDetalle.ValorBinario == null || DocumentoDetalle.ValorBinario.Length.Equals(0))
                //    baseDatosDA.AsignarParametroNulo("@PVB_VALOR_BINARIO", true, ParameterDirection.Input, DbType.Binary);
                //else
                //    baseDatosDA.AsignarParametroArrayByte("@PVB_VALOR_BINARIO", DocumentoDetalle.ValorBinario, true, ParameterDirection.Input, DbType.Binary);
                //if (DocumentoDetalle.ExtensionArchivo.Equals(""))
                //    baseDatosDA.AsignarParametroNulo("@PVC_EXTENSION_ARCHIVO", true);
                //else
                //    baseDatosDA.AsignarParametroCadena("@PVC_EXTENSION_ARCHIVO", DocumentoDetalle.ExtensionArchivo, true);
                baseDatosDA.AsignarParametroBoolean("@PBL_APROBADO", DocumentoDetalle.Aprobado, true);
                baseDatosDA.AsignarParametroCadena("@PVC_COMENTARIO", DocumentoDetalle.Comentario, true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_USUARIO_CRE", DocumentoDetalle.UsuarioCreacion.IdUsuario, true);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_INSERCION", DocumentoDetalle.TipoInsercion, true);

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

        public static void ActualizarDocumentoDetalle(DocumentoDetalleBE DocumentoDetalle, DBBaseDatos BaseDatos = null)
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
                baseDatosDA.CrearComando("USP_DOCUMENTO_DET", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "U", true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_DOCUMENTO", DocumentoDetalle.Documento.Documento.IdValor, true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_TAREA", DocumentoDetalle.Documento.Tarea.IdTarea, true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_CAMPO", DocumentoDetalle.Campo.IdValor, true);
                if (DocumentoDetalle.IdValor.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PVC_VALOR_CAMPO", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PVC_VALOR_CAMPO", DocumentoDetalle.IdValor, true);
                if (DocumentoDetalle.ValorCadena.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PVC_VALOR_CADENA", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PVC_VALOR_CADENA", DocumentoDetalle.ValorCadena, true);
                if (DocumentoDetalle.ValorEntero.Equals(null))
                    baseDatosDA.AsignarParametroNulo("@PIN_VALOR_ENTERO", true);
                else
                    baseDatosDA.AsignarParametroEntero("@PIN_VALOR_ENTERO", (Int32)DocumentoDetalle.ValorEntero, true);
                if (DocumentoDetalle.ValorNumerico.Equals(null))
                    baseDatosDA.AsignarParametroNulo("@PNU_VALOR_NUMERICO", true);
                else
                    baseDatosDA.AsignarParametroDouble("@PNU_VALOR_NUMERICO", (Double)DocumentoDetalle.ValorNumerico, true);
                if (DocumentoDetalle.ValorFecha.Equals(Convert.ToDateTime("01/01/0001 00:00:00.00")))
                    baseDatosDA.AsignarParametroNulo("@PDT_VALOR_FECHA", true);
                else
                    baseDatosDA.AsignarParametroFecha("@PDT_VALOR_FECHA", DocumentoDetalle.ValorFecha, true);
                if (DocumentoDetalle.ValorBoolean.Equals(null))
                    baseDatosDA.AsignarParametroNulo("@PBL_VALOR_BOOLEANO", true);
                else
                    baseDatosDA.AsignarParametroBoolean("@PBL_VALOR_BOOLEANO", (Boolean)DocumentoDetalle.ValorBoolean, true);
                if (DocumentoDetalle.ValorBinario == null || DocumentoDetalle.ValorBinario.Length.Equals(0))
                    baseDatosDA.AsignarParametroNulo("@PVB_VALOR_BINARIO", true, ParameterDirection.Input, DbType.Binary);
                else
                    baseDatosDA.AsignarParametroArrayByte("@PVB_VALOR_BINARIO", DocumentoDetalle.ValorBinario, true, ParameterDirection.Input, DbType.Binary);
                if (DocumentoDetalle.ExtensionArchivo.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PVC_EXTENSION_ARCHIVO", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PVC_EXTENSION_ARCHIVO", DocumentoDetalle.ExtensionArchivo, true);
                baseDatosDA.AsignarParametroBoolean("@PBL_APROBADO", DocumentoDetalle.Aprobado, true);
                baseDatosDA.AsignarParametroCadena("@PVC_COMENTARIO", DocumentoDetalle.Comentario, true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_USUARIO_MOD", DocumentoDetalle.UsuarioModificacion.IdUsuario, true);

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

        public static void EliminarFisicoEntidadDetalleProceso(DocumentoDetalleBE DocumentoDetalle, DBBaseDatos BaseDatos = null)
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
                baseDatosDA.CrearComando("USP_DOCUMENTO_DET_PROC", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "F", true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_DOCUMENTO", DocumentoDetalle.Documento.Documento.IdValor, true);
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
