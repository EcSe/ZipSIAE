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
    public class DocumentoMaterialBL
    {
        public static List<DocumentoMaterialBE> ListarDocumentoMaterial(DocumentoMaterialBE DocumentoMaterial)
        {
            List<DocumentoMaterialBE> lstResultado = new List<DocumentoMaterialBE>();
            DBBaseDatos baseDatos = new DBBaseDatos();
            baseDatos.Configurar();
            baseDatos.Conectar();
            try
            {
                baseDatos.CrearComando("USP_DOCUMENTO_MATERIAL", CommandType.StoredProcedure);
                baseDatos.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "S", true);

                if (DocumentoMaterial.Documento.Documento.IdValor.Equals(""))
                    baseDatos.AsignarParametroNulo("@PCH_ID_DOCUMENTO", true);
                else
                    baseDatos.AsignarParametroCadena("@PCH_ID_DOCUMENTO", DocumentoMaterial.Documento.Documento.IdValor, true);
                if (DocumentoMaterial.Documento.Tarea.IdTarea.Equals(""))
                    baseDatos.AsignarParametroNulo("@PCH_ID_TAREA", true);
                else
                    baseDatos.AsignarParametroCadena("@PCH_ID_TAREA", DocumentoMaterial.Documento.Tarea.IdTarea, true);
                if (DocumentoMaterial.Documento.Tarea.NodoIIBBA.IdNodo.Equals(""))
                    baseDatos.AsignarParametroNulo("@PCH_ID_NODO", true);
                else
                    baseDatos.AsignarParametroCadena("@PCH_ID_NODO", DocumentoMaterial.Documento.Tarea.NodoIIBBA.IdNodo, true);

                DbDataReader drDatos = baseDatos.EjecutarConsulta();

                while (drDatos.Read())
                {
                    DocumentoMaterialBE item = new DocumentoMaterialBE();
                    item.Documento = new DocumentoBE();

                    item.Documento.Documento.IdValor = drDatos.GetString(drDatos.GetOrdinal("CH_ID_DOCUMENTO"));
                    item.Documento.Tarea.IdTarea = drDatos.GetString(drDatos.GetOrdinal("CH_ID_TAREA"));
                    item.Documento.Tarea.NodoIIBBA.IdNodo = drDatos.GetString(drDatos.GetOrdinal("CH_ID_NODO"));
                    item.Material.Entidad.IdEntidad = drDatos.GetString(drDatos.GetOrdinal("VC_ID_ENTIDAD"));
                    //item.Cantidad = drDatos.GetInt32(drDatos.GetOrdinal("IN_CANTIDAD"));
                    item.Cantidad = (Double)drDatos.GetDecimal(drDatos.GetOrdinal("NU_CANTIDAD"));
                    item.Material.IdValor = drDatos.GetString(drDatos.GetOrdinal("VC_ID_VALOR"));
                    item.Material.ValorCadena1 = drDatos.GetString(drDatos.GetOrdinal("VC_VALOR_CADENA1"));
                    item.Material.ValorCadena2 = drDatos.GetString(drDatos.GetOrdinal("VC_VALOR_CADENA2"));

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

        public static void InsertarDocumentoMaterial(DocumentoMaterialBE DocumentoMaterial, DBBaseDatos BaseDatos = null)
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
                baseDatosDA.CrearComando("USP_DOCUMENTO_MATERIAL", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "I", true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_DOCUMENTO", DocumentoMaterial.Documento.Documento.IdValor, true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_TAREA", DocumentoMaterial.Documento.Tarea.IdTarea, true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_NODO", DocumentoMaterial.Documento.Tarea.NodoIIBBA.IdNodo, true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_MATERIAL", DocumentoMaterial.Material.IdValor, true);
                baseDatosDA.AsignarParametroDouble("@PNU_CANTIDAD", DocumentoMaterial.Cantidad, true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_USUARIO_CRE", DocumentoMaterial.UsuarioCreacion.IdUsuario, true);
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

        public static void ActualizarDocumentoMaterial(DocumentoMaterialBE DocumentoMaterial, DBBaseDatos BaseDatos = null)
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
                baseDatosDA.CrearComando("USP_DOCUMENTO_MATERIAL", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "U", true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_DOCUMENTO", DocumentoMaterial.Documento.Documento.IdValor, true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_TAREA", DocumentoMaterial.Documento.Tarea.IdTarea, true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_NODO", DocumentoMaterial.Documento.Tarea.NodoIIBBA.IdNodo, true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_MATERIAL", DocumentoMaterial.Material.IdValor, true);
                baseDatosDA.AsignarParametroDouble("@PNU_CANTIDAD", DocumentoMaterial.Cantidad, true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_USUARIO_CRE", DocumentoMaterial.UsuarioCreacion.IdUsuario, true);
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
