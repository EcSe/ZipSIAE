using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BusinessEntity;
using DataAccess;
using System.Data;

namespace BusinessLogic
{
    public class PMPBL
    {
        public static void InsertarPMPProceso(PMPBE PMP, DBBaseDatos BaseDatos = null)
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
                baseDatosDA.CrearComando("USP_PMP_PROC", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "I", true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_NODO", PMP.Nodo.IdNodo, true);

                baseDatosDA.AsignarParametroCadena("@PCH_ID_IIBB", PMP.Detalles[0].InstitucionBeneficiaria.IdInstitucionBeneficiaria, true);
                //baseDatosDA.AsignarParametroCadena("@PVC_NOMBRE_IIBB", PMP.Detalles[0].InstitucionBeneficiaria.Nombre, true);
                //baseDatosDA.AsignarParametroDouble("@PNU_LATITUD_IIBB", PMP.Detalles[0].InstitucionBeneficiaria.Latitud, true);
                //baseDatosDA.AsignarParametroDouble("@PNU_LONGITUD_IIBB", PMP.Detalles[0].InstitucionBeneficiaria.Longitud, true);

                baseDatosDA.AsignarParametroCadena("@PVC_NOM_MOD_ANT_NODO", PMP.ModeloAntenaNodo.ValorCadena1, true);
                baseDatosDA.AsignarParametroEntero("@PIN_GAIN_ANT_NODO", PMP.GananciaAntenaNodo, true);
                baseDatosDA.AsignarParametroEntero("@PIN_ALT_ANT_NODO", PMP.AlturaAntenaNodo, true);
                baseDatosDA.AsignarParametroEntero("@PIN_AZI_ANT_NODO", PMP.AzimuthAntenaNodo, true);
                baseDatosDA.AsignarParametroEntero("@PIN_ELE_ANT_NODO", PMP.ElevacionAntenaNodo, true);
                baseDatosDA.AsignarParametroEntero("@PIN_EIRP_ANT_NODO", PMP.EIRPAntenaNodo, true);

                baseDatosDA.AsignarParametroEntero("@PIN_SECTOR_IIBB", PMP.Detalles[0].SectorIIBB, true);
                baseDatosDA.AsignarParametroDouble("@PNU_AZI_ANT_IIBB", PMP.Detalles[0].AzimuthAntenaIIBB, true);
                baseDatosDA.AsignarParametroDouble("@PNU_ELE_ANT_IIBB", PMP.Detalles[0].ElevacionAntenaIIBB, true);
                baseDatosDA.AsignarParametroEntero("@PIN_TX_TORRE_IIBB", PMP.Detalles[0].TXTorreIIBB, true);
                baseDatosDA.AsignarParametroDouble("@PNU_EIRP_ANT_IIBB", PMP.Detalles[0].EIRPAntenaIIBB, true);
                baseDatosDA.AsignarParametroDouble("@PNU_NIVEL_RX_NODO", PMP.Detalles[0].NivelRXNodo, true);
                baseDatosDA.AsignarParametroDouble("@PNU_NIVEL_RX_IIBB", PMP.Detalles[0].NivelRXIIBB, true);
                baseDatosDA.AsignarParametroDouble("@PNU_FADE_MARGIN_NODO", PMP.Detalles[0].FadeMarginNodo, true);
                baseDatosDA.AsignarParametroDouble("@PNU_FADE_MARGIN_IIBB", PMP.Detalles[0].FadeMarginIIBB, true);
                baseDatosDA.AsignarParametroDouble("@PNU_DISPONIBILIDAD", PMP.Detalles[0].Disponibilidad, true);
                baseDatosDA.AsignarParametroDouble("@PNU_DISTANCIA", PMP.Detalles[0].Distancia, true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_USUARIO_CRE", PMP.UsuarioCreacion.IdUsuario, true);

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

        public static void EliminarFisicoPMPProceso(PMPBE PMP, DBBaseDatos BaseDatos = null)
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
                baseDatosDA.CrearComando("USP_PMP_PROC", CommandType.StoredProcedure);
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
