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
    public class DocumentoIPEquipamientoBL
    {
        public static void EliminarFisicoDocumentoIPEquipamientoProceso(DocumentoIPEquipamientoBE DocumentoIPEquipamiento, DBBaseDatos BaseDatos = null)
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
                baseDatosDA.CrearComando("USP_DOCUMENTO_IP_EQUIPAMIENTO_PROC", CommandType.StoredProcedure);
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

        public static void InsertarDocumentoIPEquipamientoProceso(DocumentoIPEquipamientoBE DocumentoIPEquipamiento, DBBaseDatos BaseDatos = null)
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
                baseDatosDA.CrearComando("USP_DOCUMENTO_IP_EQUIPAMIENTO_PROC", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "I", true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_DOCUMENTO", DocumentoIPEquipamiento.DocumentoIP.Documento.Documento.IdValor, true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_TAREA", DocumentoIPEquipamiento.DocumentoIP.Documento.Tarea.IdTarea,true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_NODO", DocumentoIPEquipamiento.DocumentoIP.Documento.Tarea.NodoIIBBA.IdNodo, true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_EQUIPAMIENTO", DocumentoIPEquipamiento.Equipamiento.IdValor, true);
                baseDatosDA.AsignarParametroCadena("@PVC_IP_EQUIPAMIENTO", DocumentoIPEquipamiento.IPEquipamiento, true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_USUARIO_CRE", DocumentoIPEquipamiento.UsuarioCreacion.IdUsuario, true);

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
