using System;
using System.Collections.Generic;
using BusinessEntity;
using DataAccess;
using System.Data;
using System.Data.Common;

namespace BusinessLogic
{
    public class UsuarioBL
    {
        public static List<UsuarioBE> ListarUsuarios(UsuarioBE usuarioBE)
        {
            List<UsuarioBE> lstResultadosBE = new List<UsuarioBE>();
            DBBaseDatos baseDatosDA = new DBBaseDatos();
            baseDatosDA.Configurar();
            baseDatosDA.Conectar();
            try
            {
                baseDatosDA.CrearComando("USP_USUARIO", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "S", true);
                if (usuarioBE.IdUsuario.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PVC_ID_USUARIO", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PVC_ID_USUARIO", usuarioBE.IdUsuario, true);
                if (usuarioBE.TipoDocumento.IdValor.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PCH_ID_TIP_DOC", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PCH_ID_TIP_DOC", usuarioBE.TipoDocumento.IdValor, true);
                if (usuarioBE.NumeroDocumento.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PVC_NUM_DOC", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PVC_NUM_DOC", usuarioBE.NumeroDocumento, true);
                if (usuarioBE.NombreRazon.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PVC_NOMBRE_RAZON", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PVC_NOMBRE_RAZON", usuarioBE.NombreRazon, true);
                if (usuarioBE.ApellidoPaterno.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PVC_APE_PAT", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PVC_APE_PAT", usuarioBE.ApellidoPaterno, true);
                if (usuarioBE.ApellidoMaterno.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PVC_APE_MAT", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PVC_APE_MAT", usuarioBE.ApellidoMaterno, true);
                if (usuarioBE.Password.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PVC_PASSWORD", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PVC_PASSWORD", usuarioBE.Password, true);
                if (usuarioBE.Perfil.IdValor.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PCH_ID_PERFIL", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PCH_ID_PERFIL", usuarioBE.Perfil.IdValor, true);
                //if (usuarioBE.Contratista.TipoDocumento.IdValor.Equals(""))
                //    baseDatosDA.AsignarParametroNulo("@PCH_ID_TIP_DOC_CONT", true);
                //else
                //    baseDatosDA.AsignarParametroCadena("@PCH_ID_TIP_DOC_CONT", usuarioBE.Contratista.TipoDocumento.IdValor, true);
                //if (usuarioBE.Contratista.NumeroDocumento.Equals(""))
                //    baseDatosDA.AsignarParametroNulo("@PVC_NUM_DOC_CONT", true);
                //else
                //    baseDatosDA.AsignarParametroCadena("@PVC_NUM_DOC_CONT", usuarioBE.Contratista.NumeroDocumento, true);
                if (usuarioBE.Contratista.IdValor.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PCH_ID_CONTRATISTA", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PCH_ID_CONTRATISTA", usuarioBE.Contratista.IdValor, true);
                if (usuarioBE.UsuarioCreacion == null || usuarioBE.UsuarioCreacion.Perfil.IdValor.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PVC_ID_PERFIL_U", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PVC_ID_PERFIL_U", usuarioBE.UsuarioCreacion.Perfil.IdValor, true);
                if (usuarioBE.Metodo.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PCH_ID_METODO", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PCH_ID_METODO", usuarioBE.Metodo, true);
                DbDataReader drDatos = baseDatosDA.EjecutarConsulta();

                while (drDatos.Read())
                {
                    UsuarioBE item = new UsuarioBE();

                    item.IdUsuario = drDatos.GetString(drDatos.GetOrdinal("VC_ID_USUARIO"));
                    item.TipoDocumento.IdValor = drDatos.GetString(drDatos.GetOrdinal("CH_ID_TIP_DOC"));
                    item.TipoDocumento.ValorCadena2 = drDatos.GetString(drDatos.GetOrdinal("ABREV_TIP_DOC"));
                    item.NumeroDocumento = drDatos.GetString(drDatos.GetOrdinal("VC_NUM_DOC"));
                    item.NombreRazon = drDatos.GetString(drDatos.GetOrdinal("VC_NOMBRE_RAZON"));
                    if (!drDatos.IsDBNull(drDatos.GetOrdinal("VC_APE_PAT")))
                        item.ApellidoPaterno = drDatos.GetString(drDatos.GetOrdinal("VC_APE_PAT"));
                    if (!drDatos.IsDBNull(drDatos.GetOrdinal("VC_APE_MAT")))
                        item.ApellidoMaterno = drDatos.GetString(drDatos.GetOrdinal("VC_APE_MAT"));
                    item.NombreCompleto = drDatos.GetString(drDatos.GetOrdinal("VC_NOMBRE_COMP"));
                    item.Password = drDatos.GetString(drDatos.GetOrdinal("VC_PASSWORD"));
                    item.Perfil.IdValor = drDatos.GetString(drDatos.GetOrdinal("CH_ID_PERFIL"));
                    item.Perfil.ValorCadena1 = drDatos.GetString(drDatos.GetOrdinal("VC_NOMBRE_PERFIL"));
                    //if (!drDatos.IsDBNull(drDatos.GetOrdinal("CH_ID_TIP_DOC_CONT")))
                    //    item.Contratista.TipoDocumento.IdValor = drDatos.GetString(drDatos.GetOrdinal("CH_ID_TIP_DOC_CONT"));
                    //if (!drDatos.IsDBNull(drDatos.GetOrdinal("VC_NUM_DOC_CONT")))
                    //    item.Contratista.NumeroDocumento = drDatos.GetString(drDatos.GetOrdinal("VC_NUM_DOC_CONT"));
                    //if (!drDatos.IsDBNull(drDatos.GetOrdinal("VC_NOM_COMP_CONT")))
                    //    item.Contratista.NombreCompleto = drDatos.GetString(drDatos.GetOrdinal("VC_NOM_COMP_CONT"));
                    if (!drDatos.IsDBNull(drDatos.GetOrdinal("CH_ID_CONTRATISTA")))
                        item.Contratista.IdValor = drDatos.GetString(drDatos.GetOrdinal("CH_ID_CONTRATISTA"));
                    if (!drDatos.IsDBNull(drDatos.GetOrdinal("VC_NOM_COMP_CONT")))
                        item.Contratista.ValorCadena1 = drDatos.GetString(drDatos.GetOrdinal("VC_NOM_COMP_CONT"));
                    item.Email = drDatos.GetString(drDatos.GetOrdinal("VC_EMAIL"));

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

        public static void InsertarUsuario(UsuarioBE usuarioBE)
        {
            DBBaseDatos baseDatosDA = new DBBaseDatos();
            baseDatosDA.Configurar();
            baseDatosDA.Conectar();
            try
            {
                baseDatosDA.CrearComando("USP_USUARIO", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "I", true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_USUARIO", usuarioBE.IdUsuario, true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_TIP_DOC", usuarioBE.TipoDocumento.IdValor, true);
                baseDatosDA.AsignarParametroCadena("@PVC_NUM_DOC", usuarioBE.NumeroDocumento, true);
                baseDatosDA.AsignarParametroCadena("@PVC_NOMBRE_RAZON", usuarioBE.NombreRazon, true);
                baseDatosDA.AsignarParametroCadena("@PVC_APE_PAT", usuarioBE.ApellidoPaterno, true);
                baseDatosDA.AsignarParametroCadena("@PVC_APE_MAT", usuarioBE.ApellidoMaterno, true);
                baseDatosDA.AsignarParametroCadena("@PVC_PASSWORD", usuarioBE.Password, true);
                baseDatosDA.AsignarParametroCadena("@PVC_EMAIL", usuarioBE.Email, true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_PERFIL", usuarioBE.Perfil.IdValor, true);
                //if (!usuarioBE.Contratista.TipoDocumento.IdValor.Equals(""))
                //    baseDatosDA.AsignarParametroCadena("@PCH_ID_TIP_DOC_CONT", usuarioBE.Contratista.TipoDocumento.IdValor, true);
                //if (!usuarioBE.Contratista.NumeroDocumento.Equals(""))
                //    baseDatosDA.AsignarParametroCadena("@PVC_NUM_DOC_CONT", usuarioBE.Contratista.NumeroDocumento, true);
                if (usuarioBE.Contratista.IdValor.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PCH_ID_CONTRATISTA", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PCH_ID_CONTRATISTA", usuarioBE.Contratista.IdValor, true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_USUARIO_CRE", usuarioBE.UsuarioCreacion.IdUsuario, true);

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

        public static void EditarUsuario(UsuarioBE usuarioBE)
        {
            DBBaseDatos baseDatosDA = new DBBaseDatos();
            baseDatosDA.Configurar();
            baseDatosDA.Conectar();
            try
            {
                baseDatosDA.CrearComando("USP_USUARIO", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "U", true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_USUARIO", usuarioBE.IdUsuario, true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_TIP_DOC", usuarioBE.TipoDocumento.IdValor, true);
                baseDatosDA.AsignarParametroCadena("@PVC_NUM_DOC", usuarioBE.NumeroDocumento, true);
                baseDatosDA.AsignarParametroCadena("@PVC_NOMBRE_RAZON", usuarioBE.NombreRazon, true);
                baseDatosDA.AsignarParametroCadena("@PVC_APE_PAT", usuarioBE.ApellidoPaterno, true);
                baseDatosDA.AsignarParametroCadena("@PVC_APE_MAT", usuarioBE.ApellidoMaterno, true);
                baseDatosDA.AsignarParametroCadena("@PVC_PASSWORD", usuarioBE.Password, true);
                baseDatosDA.AsignarParametroCadena("@PVC_EMAIL", usuarioBE.Email, true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_PERFIL", usuarioBE.Perfil.IdValor, true);
                //if (!usuarioBE.Contratista.TipoDocumento.IdValor.Equals(""))
                //    baseDatosDA.AsignarParametroCadena("@PCH_ID_TIP_DOC_CONT", usuarioBE.Contratista.TipoDocumento.IdValor, true);
                //if (!usuarioBE.Contratista.NumeroDocumento.Equals(""))
                //    baseDatosDA.AsignarParametroCadena("@PVC_NUM_DOC_CONT", usuarioBE.Contratista.NumeroDocumento, true);
                if (usuarioBE.Contratista.IdValor.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PCH_ID_CONTRATISTA", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PCH_ID_CONTRATISTA", usuarioBE.Contratista.IdValor, true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_USUARIO_MOD", usuarioBE.UsuarioModificacion.IdUsuario, true);

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

        public static void EliminarUsuario(UsuarioBE usuarioBE)
        {
            DBBaseDatos baseDatosDA = new DBBaseDatos();
            baseDatosDA.Configurar();
            baseDatosDA.Conectar();
            try
            {
                baseDatosDA.CrearComando("USP_USUARIO", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "D", true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_USUARIO", usuarioBE.IdUsuario, true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_USUARIO_MOD", usuarioBE.UsuarioModificacion.IdUsuario, true);

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

        public static void GenerarTicket(UsuarioBE usuarioBE)
        {
            DBBaseDatos baseDatosDA = new DBBaseDatos();
            baseDatosDA.Configurar();
            baseDatosDA.Conectar();
            try
            {
                baseDatosDA.CrearComando("USP_GENERAR_TICKET", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_USUARIO", usuarioBE.IdUsuario, true);
                baseDatosDA.AsignarParametroCadena("@PVC_TICKET", "", 20, true, ParameterDirection.Output);

                baseDatosDA.EjecutarComando();
                usuarioBE.Ticket = baseDatosDA.DevolverParametroCadena("@PVC_TICKET");
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

        public static void ObtenerUsuarioTicket(UsuarioBE usuarioBE)
        {
            DBBaseDatos baseDatosDA = new DBBaseDatos();
            baseDatosDA.Configurar();
            baseDatosDA.Conectar();
            try
            {
                baseDatosDA.CrearComando("USP_OBTENER_USUARIO_TICKET", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PVC_TICKET", usuarioBE.Ticket, true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_USUARIO", "", 20, true, ParameterDirection.Output);
                baseDatosDA.AsignarParametroCadena("@PVC_PASSWORD", "", 50, true, ParameterDirection.Output);

                baseDatosDA.EjecutarComando();

                usuarioBE.IdUsuario = baseDatosDA.DevolverParametroCadena("@PVC_ID_USUARIO");
                usuarioBE.Password = baseDatosDA.DevolverParametroCadena("@PVC_PASSWORD");
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
