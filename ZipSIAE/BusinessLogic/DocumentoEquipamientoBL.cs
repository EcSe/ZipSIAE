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
    public class DocumentoEquipamientoBL
    {
        public static List<DocumentoEquipamientoBE> ListarDocumentoEquipamiento(DocumentoEquipamientoBE DocumentoEquipamiento)
        {
            List<DocumentoEquipamientoBE> lstResultado = new List<DocumentoEquipamientoBE>();
            DBBaseDatos baseDatos = new DBBaseDatos();
            baseDatos.Configurar();
            baseDatos.Conectar();
            try
            {
                baseDatos.CrearComando("USP_DOCUMENTO_EQUIPAMIENTO", CommandType.StoredProcedure);
                baseDatos.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "S", true);

                if (DocumentoEquipamiento.Documento.Documento.IdValor.Equals(""))
                    baseDatos.AsignarParametroNulo("@PCH_ID_DOCUMENTO", true);
                else
                    baseDatos.AsignarParametroCadena("@PCH_ID_DOCUMENTO", DocumentoEquipamiento.Documento.Documento.IdValor, true);
                if (DocumentoEquipamiento.Documento.Tarea.IdTarea.Equals(""))
                    baseDatos.AsignarParametroNulo("@PCH_ID_TAREA", true);
                else
                    baseDatos.AsignarParametroCadena("@PCH_ID_TAREA", DocumentoEquipamiento.Documento.Tarea.IdTarea, true);
                if (DocumentoEquipamiento.Documento.Tarea.NodoIIBBA.IdNodo.Equals(""))
                    baseDatos.AsignarParametroNulo("@PVC_ID_NODO_IIBB", true);
                else
                    baseDatos.AsignarParametroCadena("@PVC_ID_NODO_IIBB", DocumentoEquipamiento.Documento.Tarea.NodoIIBBA.IdNodo, true);

                DbDataReader drDatos = baseDatos.EjecutarConsulta();

                while (drDatos.Read())
                {
                    DocumentoEquipamientoBE item = new DocumentoEquipamientoBE();
                    item.Documento = new DocumentoBE();

                    item.Documento.Documento.IdValor = drDatos.GetString(drDatos.GetOrdinal("CH_ID_DOCUMENTO"));
                    item.Documento.Tarea.IdTarea = drDatos.GetString(drDatos.GetOrdinal("CH_ID_TAREA"));
                    item.Documento.Tarea.NodoIIBBA.IdNodo = drDatos.GetString(drDatos.GetOrdinal("VC_ID_NODO_IIBB"));
                    item.Equipamiento.Entidad.IdEntidad = drDatos.GetString(drDatos.GetOrdinal("VC_ID_ENTIDAD"));
                    item.Item = drDatos.GetInt32(drDatos.GetOrdinal("IN_ITEM"));
                    item.SerieEquipamiento = drDatos.GetString(drDatos.GetOrdinal("VC_SERIE_EQUIPAMIENTO"));
                    item.Equipamiento.IdValor = drDatos.GetString(drDatos.GetOrdinal("VC_ID_VALOR"));
                    item.Equipamiento.ValorCadena1 = drDatos.GetString(drDatos.GetOrdinal("VC_VALOR_CADENA1"));
                    if (!drDatos.IsDBNull(drDatos.GetOrdinal("VC_VALOR_CADENA2")))
                        item.Equipamiento.ValorCadena2 = drDatos.GetString(drDatos.GetOrdinal("VC_VALOR_CADENA2"));
                    if (!drDatos.IsDBNull(drDatos.GetOrdinal("VC_VALOR_CADENA3")))
                        item.Equipamiento.ValorCadena3 = drDatos.GetString(drDatos.GetOrdinal("VC_VALOR_CADENA3"));
                    if (!drDatos.IsDBNull(drDatos.GetOrdinal("VC_VALOR_CADENA4")))
                        item.Equipamiento.ValorCadena4 = drDatos.GetString(drDatos.GetOrdinal("VC_VALOR_CADENA4"));

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
                baseDatos.Desconectar();
                baseDatos = null;
            }

            return lstResultado;
        }

        public static void InsertarDocumentoEquipamiento(DocumentoEquipamientoBE DocumentoEquipamiento,
            DBBaseDatos BaseDatos = null)
        {
            DBBaseDatos baseDatosDA = new DBBaseDatos();
            if (BaseDatos == null)
            {
                baseDatosDA.Configurar();
                baseDatosDA.Conectar();
            }
            else
                baseDatosDA = BaseDatos;

            try
            {
                baseDatosDA.CrearComando("USP_DOCUMENTO_EQUIPAMIENTO", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "I", true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_DOCUMENTO", DocumentoEquipamiento.Documento.Documento.IdValor, true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_TAREA", DocumentoEquipamiento.Documento.Tarea.IdTarea, true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_NODO", DocumentoEquipamiento.Documento.Tarea.NodoIIBBA.IdNodo, true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_EQUIPAMIENTO", DocumentoEquipamiento.Equipamiento.IdValor, true);
                baseDatosDA.AsignarParametroCadena("@PVC_SERIE_EQUIPAMIENTO", DocumentoEquipamiento.SerieEquipamiento, true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_USUARIO_CRE", DocumentoEquipamiento.UsuarioCreacion.IdUsuario, true);
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

        public static void ActualizarDocumentoEquipamiento(DocumentoEquipamientoBE DocumentoEquipamiento,
            DBBaseDatos BaseDatos = null)
        {
            DBBaseDatos baseDatosDA = new DBBaseDatos();
            if (BaseDatos == null)
            {
                baseDatosDA.Configurar();
                baseDatosDA.Conectar();
            }
            else
                baseDatosDA = BaseDatos;

            try
            {
                baseDatosDA.CrearComando("USP_DOCUMENTO_EQUIPAMIENTO", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "U", true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_DOCUMENTO", DocumentoEquipamiento.Documento.Documento.IdValor, true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_TAREA", DocumentoEquipamiento.Documento.Tarea.IdTarea, true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_NODO", DocumentoEquipamiento.Documento.Tarea.NodoIIBBA.IdNodo, true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_EQUIPAMIENTO", DocumentoEquipamiento.Equipamiento.IdValor, true);
                baseDatosDA.AsignarParametroCadena("@PVC_SERIE_EQUIPAMIENTO", DocumentoEquipamiento.SerieEquipamiento, true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_USUARIO_CRE", DocumentoEquipamiento.UsuarioCreacion.IdUsuario, true);
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

        public static void InsertarDocumentoEquipamientoProceso(DocumentoEquipamientoBE DocumentoEquipamiento, DBBaseDatos BaseDatos = null)
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
                baseDatosDA.CrearComando("USP_DOCUMENTO_EQUIPAMIENTO_PROC", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "I", true);
                baseDatosDA.AsignarParametroCadena("@PVC_EMPRESA_EQUIPAMIENTO", DocumentoEquipamiento.IdEmpresa, true);
                if (DocumentoEquipamiento.Documento.Tarea.IdTarea == null)
                    baseDatosDA.AsignarParametroNulo("@PCH_ID_TAREA", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PCH_ID_TAREA", DocumentoEquipamiento.Documento.Tarea.IdTarea, true);
                if (DocumentoEquipamiento.Documento.Tarea.NodoIIBBA.IdNodo == null)
                    baseDatosDA.AsignarParametroNulo("@PVC_ID_NODO_IIBB", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PVC_ID_NODO_IIBB", DocumentoEquipamiento.Documento.Tarea.NodoIIBBA.IdNodo, true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_EQUIPAMIENTO", DocumentoEquipamiento.Equipamiento.IdValor, true);
                baseDatosDA.AsignarParametroCadena("@PVC_SERIE_EQUIPAMIENTO", DocumentoEquipamiento.SerieEquipamiento, true);
                baseDatosDA.AsignarParametroCadena("@PVC_SERIE_KIT", DocumentoEquipamiento.KitSIAE.SerieKit, true);
                baseDatosDA.AsignarParametroCadena("@PVC_CODIGO_GILAT", DocumentoEquipamiento.KitSIAE.CodigoGilat, true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_USUARIO_CRE", DocumentoEquipamiento.UsuarioCreacion.IdUsuario, true);

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

        public static void InsertarDocumentoEquipamientoAlimentacionProceso(DocumentoEquipamientoBE DocumentoEquipamiento, DBBaseDatos BaseDatos = null)
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
                baseDatosDA.CrearComando("USP_DOCUMENTO_EQUIPAMIENTO_PROC", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "A", true);
                baseDatosDA.AsignarParametroCadena("@PVC_EMPRESA_EQUIPAMIENTO", DocumentoEquipamiento.IdEmpresa, true);
                if (DocumentoEquipamiento.Documento.Tarea.IdTarea == null)
                    baseDatosDA.AsignarParametroNulo("@PCH_ID_TAREA", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PCH_ID_TAREA", DocumentoEquipamiento.Documento.Tarea.IdTarea, true);
                if (DocumentoEquipamiento.Documento.Tarea.NodoIIBBA.IdNodo == null)
                    baseDatosDA.AsignarParametroNulo("@PVC_ID_NODO_IIBB", true);
                else
                    baseDatosDA.AsignarParametroCadena("@PVC_ID_NODO_IIBB", DocumentoEquipamiento.Documento.Tarea.NodoIIBBA.IdNodo, true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_EQUIPAMIENTO", DocumentoEquipamiento.Equipamiento.IdValor, true);
                baseDatosDA.AsignarParametroCadena("@PVC_SERIE_EQUIPAMIENTO", DocumentoEquipamiento.SerieEquipamiento, true);
                baseDatosDA.AsignarParametroCadena("@PVC_SERIE_KIT", DocumentoEquipamiento.KitSIAE.SerieKit, true);
                baseDatosDA.AsignarParametroCadena("@PVC_CODIGO_GILAT", DocumentoEquipamiento.KitSIAE.CodigoGilat, true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_USUARIO_CRE", DocumentoEquipamiento.UsuarioCreacion.IdUsuario, true);

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

        public static void EliminarFisicoDocumentoEquipamientoProceso(DocumentoEquipamientoBE DocumentoEquipamiento, DBBaseDatos BaseDatos = null)
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
                baseDatosDA.CrearComando("USP_DOCUMENTO_EQUIPAMIENTO_PROC", CommandType.StoredProcedure);
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