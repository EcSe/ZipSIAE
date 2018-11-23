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
    public class TareaBL
    {
        public static List<TareaBE> ListarTareas(TareaBE Tarea,String TipoTransaccion = "S", DBBaseDatos BaseDatos = null)
        {
            List<TareaBE> lstResultadosBE = new List<TareaBE>();
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
                baseDatosDA.CrearComando("USP_TAREA", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", TipoTransaccion, true);

                if (Tarea.IdTarea.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PCH_ID_TAREA", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PCH_ID_TAREA", Tarea.IdTarea, true);

                if (!Tarea.NodoIIBBA.IdNodo.Equals(""))
                    baseDatosDA.AsignarParametroCadena("@PVC_ID_NODO_IIBB_A", Tarea.NodoIIBBA.IdNodo, true);
                else
                    baseDatosDA.AsignarParametroNulo("@PVC_ID_NODO_IIBB_A", true);
                if (Tarea.Contratista.IdValor.Equals(""))
                    baseDatosDA.AsignarParametroNulo("@PCH_ID_CONTRATISTA", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PCH_ID_CONTRATISTA", Tarea.Contratista.IdValor, true);

                DbDataReader drDatos = baseDatosDA.EjecutarConsulta();

                while (drDatos.Read())
                {
                    TareaBE item = new TareaBE();

                    if (TipoTransaccion.Equals("S"))
                    {
                        item.IdTarea = drDatos.GetString(drDatos.GetOrdinal("CH_ID_TAREA")); ;
                        item.TipoTarea.IdValor = drDatos.GetString(drDatos.GetOrdinal("CH_ID_TIP_TAREA")); ;
                        item.NodoIIBBA.IdNodo = drDatos.GetString(drDatos.GetOrdinal("VC_ID_NODO_IIBB_A")); ;
                        if (!drDatos.IsDBNull(drDatos.GetOrdinal("IN_SECTOR")))
                            item.Sector = drDatos.GetInt32(drDatos.GetOrdinal("IN_SECTOR"));
                        item.TipoTarea.IdValor = drDatos.GetString(drDatos.GetOrdinal("CH_ID_TIP_TAREA"));
                        item.TipoTarea.ValorCadena1 = drDatos.GetString(drDatos.GetOrdinal("VC_NOM_TIP_TAREA"));
                        if (!drDatos.IsDBNull(drDatos.GetOrdinal("CH_ID_NODO_B")))
                            item.NodoB.IdNodo = drDatos.GetString(drDatos.GetOrdinal("CH_ID_NODO_B"));
                        if (!drDatos.IsDBNull(drDatos.GetOrdinal("CH_ID_CONTRATISTA")))
                            item.Contratista.IdValor = drDatos.GetString(drDatos.GetOrdinal("CH_ID_CONTRATISTA"));
                    }
                    if (TipoTransaccion.Equals("Z"))
                    {
                        item.IdTarea = drDatos.GetString(drDatos.GetOrdinal("CH_ID_TAREA"));
                        //if (!drDatos.IsDBNull(drDatos.GetOrdinal("IN_SECTOR")))
                        //    item.Sector = drDatos.GetInt32(drDatos.GetOrdinal("IN_SECTOR"));
                        if (!drDatos.IsDBNull(drDatos.GetOrdinal("IN_SECTOR")))
                            item.IdSectorAP = drDatos.GetInt32(drDatos.GetOrdinal("IN_SECTOR")).ToString();
                        item.TipoTarea.IdValor = drDatos.GetString(drDatos.GetOrdinal("CH_ID_TIP_TAREA"));
                        item.TipoTarea.ValorCadena1 = drDatos.GetString(drDatos.GetOrdinal("VC_NOM_TIP_TAREA"));
                        if (!drDatos.IsDBNull(drDatos.GetOrdinal("CH_ID_TIP_NODO_A")))
                            item.TipoNodoA.IdValor = drDatos.GetString(drDatos.GetOrdinal("CH_ID_TIP_NODO_A"));
                        if (!drDatos.IsDBNull(drDatos.GetOrdinal("VC_NOM_TIP_NODO_A")))
                            item.TipoNodoA.ValorCadena1 = drDatos.GetString(drDatos.GetOrdinal("VC_NOM_TIP_NODO_A"));
                        if (!drDatos.IsDBNull(drDatos.GetOrdinal("VC_ID_NODO_IIBB_A")))
                            item.NodoIIBBA.IdNodo = drDatos.GetString(drDatos.GetOrdinal("VC_ID_NODO_IIBB_A"));
                        if (!drDatos.IsDBNull(drDatos.GetOrdinal("CH_ID_TIP_NODO_B")))
                            item.TipoNodoB.IdValor = drDatos.GetString(drDatos.GetOrdinal("CH_ID_TIP_NODO_B"));
                        if (!drDatos.IsDBNull(drDatos.GetOrdinal("VC_NOM_TIP_NODO_B")))
                            item.TipoNodoB.ValorCadena1 = drDatos.GetString(drDatos.GetOrdinal("VC_NOM_TIP_NODO_B"));
                        if (!drDatos.IsDBNull(drDatos.GetOrdinal("CH_ID_NODO_B")))
                            item.NodoB.IdNodo = drDatos.GetString(drDatos.GetOrdinal("CH_ID_NODO_B"));
                        if (!drDatos.IsDBNull(drDatos.GetOrdinal("CH_ID_CONTRATISTA")))
                            item.Contratista.IdValor = drDatos.GetString(drDatos.GetOrdinal("CH_ID_CONTRATISTA"));
                        if (!drDatos.IsDBNull(drDatos.GetOrdinal("VC_NOM_CONTRATISTA")))
                            item.Contratista.ValorCadena1 = drDatos.GetString(drDatos.GetOrdinal("VC_NOM_CONTRATISTA"));
                    }

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
                if (BaseDatos == null)
                {
                    baseDatosDA.Desconectar();
                    baseDatosDA = null;
                }
            }

            return lstResultadosBE;
        }

        public static void InsertarTareaProceso(TareaBE Tarea, DBBaseDatos BaseDatos = null)
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
                baseDatosDA.CrearComando("USP_TAREA_PROC", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "I", true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_TAREA", Tarea.IdTarea, true);
                baseDatosDA.AsignarParametroCadena("@PVC_NOMBRE_TIP_TAREA", Tarea.TipoTarea.ValorCadena1, true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_ISO_NODO", Tarea.IdIsoNodo, true);
                //if (Tarea.Contratista.NombreCompleto!=null && !Tarea.Contratista.NombreCompleto.Equals(""))
                //    baseDatosDA.AsignarParametroCadena("@PVC_NOM_COMP_CONT", Tarea.Contratista.NombreCompleto, true);
                //else
                //    baseDatosDA.AsignarParametroNulo("@PVC_NOM_COMP_CONT", true);
                if (Tarea.Contratista.ValorCadena2 != null && !Tarea.Contratista.ValorCadena2.Equals(""))
                    baseDatosDA.AsignarParametroCadena("@PCH_COD_CONT", Tarea.Contratista.ValorCadena2, true);
                else
                    baseDatosDA.AsignarParametroNulo("@PCH_COD_CONT", true);
                if (Tarea.InicioInstalacion.Equals(Convert.ToDateTime("01/01/0001 00:00:00.00")))
                    baseDatosDA.AsignarParametroNulo("@PDT_FEC_INI_INSTALACION", true);
                else
                    baseDatosDA.AsignarParametroFecha("@PDT_FEC_INI_INSTALACION", Tarea.InicioInstalacion, true);
                if (Tarea.FinInstalacion.Equals(Convert.ToDateTime("01/01/0001 00:00:00.00")))
                    baseDatosDA.AsignarParametroNulo("@PDT_FEC_FIN_INSTALACION", true);
                else
                    baseDatosDA.AsignarParametroFecha("@PDT_FEC_FIN_INSTALACION", Tarea.FinInstalacion, true);
                baseDatosDA.AsignarParametroCadena("@PVC_NOMBRE_PROY", Tarea.Proyecto.ValorCadena1, true);

                if (Tarea.TipoNodoA.ValorCadena1 != null && !Tarea.TipoNodoA.ValorCadena1.Equals(""))
                    baseDatosDA.AsignarParametroCadena("@PVC_NOMBRE_TIP_NODO_A", Tarea.TipoNodoA.ValorCadena1, true);
                else
                    baseDatosDA.AsignarParametroNulo("@PVC_NOMBRE_TIP_NODO_A", true);
                //if (Tarea.NodoIIBBA.IdNodo != null && !Tarea.NodoIIBBA.IdNodo.Equals(""))
                if (!Tarea.NodoIIBBA.IdNodo.Equals(""))
                    baseDatosDA.AsignarParametroCadena("@PVC_ID_NODO_IIBB_A", Tarea.NodoIIBBA.IdNodo, true);
                else
                    baseDatosDA.AsignarParametroNulo("@PVC_ID_NODO_IIBB_A", true);

                if (Tarea.TipoNodoB.ValorCadena1 != null && !Tarea.TipoNodoB.ValorCadena1.Equals(""))
                    baseDatosDA.AsignarParametroCadena("@PVC_NOMBRE_TIP_NODO_B", Tarea.TipoNodoB.ValorCadena1, true);
                else
                    baseDatosDA.AsignarParametroNulo("@PVC_NOMBRE_TIP_NODO_B", true);
                //if (Tarea.NodoB.IdNodo != null && !Tarea.NodoB.IdNodo.Equals(""))
                if (!Tarea.NodoB.IdNodo.Equals(""))
                    baseDatosDA.AsignarParametroCadena("@PCH_ID_NODO_B", Tarea.NodoB.IdNodo, true);
                else
                    baseDatosDA.AsignarParametroNulo("@PCH_ID_NODO_B", true);

                if (Tarea.IdSectorAP != null && !Tarea.IdSectorAP.Equals(""))
                    baseDatosDA.AsignarParametroCadena("@PVC_ID_SECTOR_AP", Tarea.IdSectorAP, true);
                else
                    baseDatosDA.AsignarParametroNulo("@PVC_ID_SECTOR_AP", true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_USUARIO_CRE", Tarea.UsuarioCreacion.IdUsuario, true);


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

        public static void EliminarFisicoTareaProceso(TareaBE Tarea, DBBaseDatos BaseDatos = null)
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
                baseDatosDA.CrearComando("USP_TAREA_PROC", CommandType.StoredProcedure);
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
