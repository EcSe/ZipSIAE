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
    public class DocumentoMedicionEnlacePropagacionBL
    {
        public static List<DocumentoMedicionEnlacePropagacionBE> ListarDocumentoMedicionEnlacePropagacion(DocumentoMedicionEnlacePropagacionBE DocumentoMedicionEnlacePropagacion)
        {
            List<DocumentoMedicionEnlacePropagacionBE> lstResultado = new List<DocumentoMedicionEnlacePropagacionBE>();
            DBBaseDatos baseDatos = new DBBaseDatos();
            baseDatos.Configurar();
            baseDatos.Conectar();
            try
            {
                baseDatos.CrearComando("USP_DOCUMENTO_MED_ENLA_PROP", CommandType.StoredProcedure);
                baseDatos.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "S", true);

                if (DocumentoMedicionEnlacePropagacion.Documento.Documento.IdValor.Equals(""))
                    baseDatos.AsignarParametroNulo("@PCH_ID_DOCUMENTO", true);
                else
                    baseDatos.AsignarParametroCadena("@PCH_ID_DOCUMENTO", DocumentoMedicionEnlacePropagacion.Documento.Documento.IdValor, true);
                if (DocumentoMedicionEnlacePropagacion.Documento.Tarea.IdTarea.Equals(""))
                    baseDatos.AsignarParametroNulo("@PCH_ID_TAREA", true);
                else
                    baseDatos.AsignarParametroCadena("@PCH_ID_TAREA", DocumentoMedicionEnlacePropagacion.Documento.Tarea.IdTarea, true);
                if (DocumentoMedicionEnlacePropagacion.NodoA.IdNodo.Equals(""))
                    baseDatos.AsignarParametroNulo("@PCH_ID_NODO_A", true);
                else
                    baseDatos.AsignarParametroCadena("@PCH_ID_NODO_A", DocumentoMedicionEnlacePropagacion.NodoA.IdNodo, true);

                DbDataReader drDatos = baseDatos.EjecutarConsulta();

                while (drDatos.Read())
                {
                    DocumentoMedicionEnlacePropagacionBE item = new DocumentoMedicionEnlacePropagacionBE();
                    item.Documento = new DocumentoBE();

                    item.Documento.Documento.IdValor = drDatos.GetString(drDatos.GetOrdinal("CH_ID_DOCUMENTO"));
                    item.Documento.Tarea.IdTarea = drDatos.GetString(drDatos.GetOrdinal("CH_ID_TAREA"));
                    item.NodoA.IdNodo = drDatos.GetString(drDatos.GetOrdinal("CH_ID_NODO_A"));
                    item.NodoIIBBB.IdNodo = drDatos.GetString(drDatos.GetOrdinal("VC_ID_NODO_IIBB_B"));
                    if (!drDatos.IsDBNull(drDatos.GetOrdinal("NU_RSS_LOCAL")))
                        item.RSSLocal = (Double)drDatos.GetDecimal(drDatos.GetOrdinal("NU_RSS_LOCAL"));
                    if (!drDatos.IsDBNull(drDatos.GetOrdinal("NU_RSS_REMOTO")))
                        item.RSSRemoto = (Double)drDatos.GetDecimal(drDatos.GetOrdinal("NU_RSS_REMOTO"));
                    if (!drDatos.IsDBNull(drDatos.GetOrdinal("NU_TIEMPO_PROMEDIO")))
                        item.TiempoPromedio = (Double)drDatos.GetDecimal(drDatos.GetOrdinal("NU_TIEMPO_PROMEDIO"));
                    if (!drDatos.IsDBNull(drDatos.GetOrdinal("NU_CAPACIDAD_SUBIDA")))
                        item.CapidadSubida = (Double)drDatos.GetDecimal(drDatos.GetOrdinal("NU_CAPACIDAD_SUBIDA"));
                    if (!drDatos.IsDBNull(drDatos.GetOrdinal("NU_CAPACIDAD_BAJADA")))
                        item.CapidadBajada = (Double)drDatos.GetDecimal(drDatos.GetOrdinal("NU_CAPACIDAD_BAJADA"));

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

        public static void InsertarDocumentoMedicionEnlacePropagacion(DocumentoMedicionEnlacePropagacionBE DocumentoMedicionEnlacePropagacion, DBBaseDatos BaseDatos = null)
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
                baseDatosDA.CrearComando("USP_DOCUMENTO_MED_ENLA_PROP", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "I", true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_DOCUMENTO", DocumentoMedicionEnlacePropagacion.Documento.Documento.IdValor, true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_TAREA", DocumentoMedicionEnlacePropagacion.Documento.Tarea.IdTarea, true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_NODO_A", DocumentoMedicionEnlacePropagacion.NodoA.IdNodo, true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_NODO_IIBB_B", DocumentoMedicionEnlacePropagacion.NodoIIBBB.IdNodo, true);
                baseDatosDA.AsignarParametroDouble("@PNU_RSS_LOCAL", DocumentoMedicionEnlacePropagacion.RSSLocal, true);
                baseDatosDA.AsignarParametroDouble("@PNU_RSS_REMOTO", DocumentoMedicionEnlacePropagacion.RSSRemoto, true);
                baseDatosDA.AsignarParametroDouble("@PNU_TIEMPO_PROMEDIO", DocumentoMedicionEnlacePropagacion.TiempoPromedio, true);
                baseDatosDA.AsignarParametroDouble("@PNU_CAPACIDAD_SUBIDA", DocumentoMedicionEnlacePropagacion.CapidadSubida, true);
                baseDatosDA.AsignarParametroDouble("@PNU_CAPACIDAD_BAJADA", DocumentoMedicionEnlacePropagacion.CapidadBajada, true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_USUARIO_CRE", DocumentoMedicionEnlacePropagacion.UsuarioCreacion.IdUsuario, true);
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

        public static void ActualizarDocumentoMedicionEnlacePropagacion(DocumentoMedicionEnlacePropagacionBE DocumentoMedicionEnlacePropagacion, DBBaseDatos BaseDatos = null)
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
                baseDatosDA.CrearComando("USP_DOCUMENTO_MED_ENLA_PROP", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "U", true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_DOCUMENTO", DocumentoMedicionEnlacePropagacion.Documento.Documento.IdValor, true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_TAREA", DocumentoMedicionEnlacePropagacion.Documento.Tarea.IdTarea, true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_NODO_A", DocumentoMedicionEnlacePropagacion.NodoA.IdNodo, true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_NODO_IIBB_B", DocumentoMedicionEnlacePropagacion.NodoIIBBB.IdNodo, true);
                baseDatosDA.AsignarParametroDouble("@PNU_RSS_LOCAL", DocumentoMedicionEnlacePropagacion.RSSLocal, true);
                baseDatosDA.AsignarParametroDouble("@PNU_RSS_REMOTO", DocumentoMedicionEnlacePropagacion.RSSRemoto, true);
                baseDatosDA.AsignarParametroDouble("@PNU_TIEMPO_PROMEDIO", DocumentoMedicionEnlacePropagacion.TiempoPromedio, true);
                baseDatosDA.AsignarParametroDouble("@PNU_CAPACIDAD_SUBIDA", DocumentoMedicionEnlacePropagacion.CapidadSubida, true);
                baseDatosDA.AsignarParametroDouble("@PNU_CAPACIDAD_BAJADA", DocumentoMedicionEnlacePropagacion.CapidadBajada, true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_USUARIO_CRE", DocumentoMedicionEnlacePropagacion.UsuarioCreacion.IdUsuario, true);
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
