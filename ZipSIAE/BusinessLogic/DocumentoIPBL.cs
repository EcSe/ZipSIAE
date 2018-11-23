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
    public class DocumentoIPBL
    {
        public static void EliminarFisicoDocumentoIPProceso(DocumentoIPBE DocumentoIP, DBBaseDatos BaseDatos = null)
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
                baseDatosDA.CrearComando("USP_DOCUMENTO_IP_PROC", CommandType.StoredProcedure);
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

        public static void InsertarDocumentoIPProceso(DocumentoIPBE DocumentoIP, DBBaseDatos BaseDatos = null)
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
                baseDatosDA.CrearComando("USP_DOCUMENTO_IP_PROC", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "I", true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_DOCUMENTO", DocumentoIP.Documento.Documento.IdValor, true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_TAREA","",7,true,ParameterDirection.Output);
                baseDatosDA.AsignarParametroCadena("@PCH_NOMBRE_TIP_NODO", DocumentoIP.Documento.Tarea.TipoNodoA.ValorCadena1, true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_NODO", DocumentoIP.Documento.Tarea.NodoIIBBA.IdNodo, true);
                if (DocumentoIP.IPSystem != null)
                    baseDatosDA.AsignarParametroCadena("@PVC_IP_SYSTEM", DocumentoIP.IPSystem, true);
                else
                    baseDatosDA.AsignarParametroNulo("@PVC_IP_SYSTEM",true);
                if (DocumentoIP.RangoGestionSeguridadEnergia!=null)
                    baseDatosDA.AsignarParametroCadena("@PVC_RANG_GEST_SEG_ENER", DocumentoIP.RangoGestionSeguridadEnergia, true);
                else
                    baseDatosDA.AsignarParametroNulo("@PVC_RANG_GEST_SEG_ENER", true);
                baseDatosDA.AsignarParametroCadena("@PVC_GATEWAY", DocumentoIP.Gateway, true);
                baseDatosDA.AsignarParametroCadena("@PVC_MASCARA", DocumentoIP.Mascara, true);
                baseDatosDA.AsignarParametroCadena("@PVC_IP_RESERVADA", DocumentoIP.IPReservada, true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_USUARIO_CRE", DocumentoIP.UsuarioCreacion.IdUsuario, true);

                baseDatosDA.EjecutarComando();

                DocumentoIP.Documento.Tarea.IdTarea = baseDatosDA.DevolverParametroCadena("@PCH_ID_TAREA");

                #region Insertamos los Equipamientos
                foreach (DocumentoIPEquipamientoBE item in DocumentoIP.Equipamientos)
                {
                    item.DocumentoIP = DocumentoIP;
                    DocumentoIPEquipamientoBL.InsertarDocumentoIPEquipamientoProceso(item, baseDatosDA);
                }
                #endregion

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
