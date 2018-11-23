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
    public class PTPBL
    {
        public static void InsertarPTPProceso(PTPBE PTP, DBBaseDatos BaseDatos = null)
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
                baseDatosDA.CrearComando("USP_PTP_PROC", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@PCH_TIPO_TRANSACCION", "I", true);
                baseDatosDA.AsignarParametroCadena("@PCH_ID_NODO_A", PTP.NodoA.IdNodo, true);
                baseDatosDA.AsignarParametroDouble("@PNU_COTA_NODO_A", PTP.CotaNodoA, true);

                baseDatosDA.AsignarParametroCadena("@PCH_ID_NODO_B", PTP.Detalles[0].NodoB.IdNodo, true);
                baseDatosDA.AsignarParametroDouble("@PNU_AZIMUTH_NODO_A", PTP.Detalles[0].AzimuthNodoA, true);
                baseDatosDA.AsignarParametroDouble("@PNU_AZIMUTH_NODO_B", PTP.Detalles[0].AzimuthNodoB, true);
                baseDatosDA.AsignarParametroDouble("@PNU_ELEVACION_NODO_A", PTP.Detalles[0].ElevacionNodoA, true);
                baseDatosDA.AsignarParametroDouble("@PNU_ELEVACION_NODO_B", PTP.Detalles[0].ElevacionNodoB, true);
                baseDatosDA.AsignarParametroDouble("@PNU_DISTANCIA", PTP.Detalles[0].Distancia, true);
                baseDatosDA.AsignarParametroDouble("@PNU_COTA_NODO_B", PTP.Detalles[0].CotaNodoB, true);
                baseDatosDA.AsignarParametroCadena("@PVC_TR_MOD_ANT_NODO_A", PTP.Detalles[0].ModeloAntenaNodoA, true);
                baseDatosDA.AsignarParametroCadena("@PVC_TR_MOD_ANT_NODO_B", PTP.Detalles[0].ModeloAntenaNodoB, true);
                baseDatosDA.AsignarParametroCadena("@PVC_TR_DIA_ANT_NODO_A", PTP.Detalles[0].DiametroAntenaNodoA, true);
                baseDatosDA.AsignarParametroCadena("@PVC_TR_DIA_ANT_NODO_B", PTP.Detalles[0].DiametroAntenaNodoB, true);
                baseDatosDA.AsignarParametroEntero("@PIN_TR_ALT_ANT_NODO_A", PTP.Detalles[0].AlturaAntenaNodoA, true);
                baseDatosDA.AsignarParametroEntero("@PIN_TR_ALT_ANT_NODO_B", PTP.Detalles[0].AlturaAntenaNodoB, true);
                baseDatosDA.AsignarParametroDouble("@PNU_TR_GAIN_ANT_NODO_A", PTP.Detalles[0].GananciaAntenaNodoA, true);
                baseDatosDA.AsignarParametroDouble("@PNU_TR_GAIN_ANT_NODO_B", PTP.Detalles[0].GananciaAntenaNodoB, true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_CANAL_NODO_A", PTP.Detalles[0].IdCanalNodoA, true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_CANAL_NODO_B", PTP.Detalles[0].IdCanalNodoB, true);
                baseDatosDA.AsignarParametroCadena("@PVC_DIS_FREC_NODO_A", PTP.Detalles[0].DisenoFrecuenciaNodoA, true);
                baseDatosDA.AsignarParametroCadena("@PVC_DIS_FREC_NODO_B", PTP.Detalles[0].DisenoFrecuenciaNodoB, true);
                baseDatosDA.AsignarParametroCadena("@PVC_NOMBRE_POLARIZACION", PTP.Detalles[0].Polarizacion.ValorCadena1, true);
                baseDatosDA.AsignarParametroCadena("@PVC_MOD_RAD_NODO_A", PTP.Detalles[0].ModeloRadioNodoA, true);
                baseDatosDA.AsignarParametroCadena("@PVC_DES_EMI_NODO_A", PTP.Detalles[0].DesignadorEmisionNodoA, true);
                baseDatosDA.AsignarParametroEntero("@PIN_POT_TX_NODO_A", PTP.Detalles[0].PotenciaTorreNodoA, true);
                baseDatosDA.AsignarParametroEntero("@PIN_POT_TX_NODO_B", PTP.Detalles[0].PotenciaTorreNodoB, true);
                baseDatosDA.AsignarParametroDouble("@PNU_EIRP_NODO_A", PTP.Detalles[0].EIRPNodoA, true);
                baseDatosDA.AsignarParametroDouble("@PNU_EIRP_NODO_B", PTP.Detalles[0].EIRPNodoB, true);
                baseDatosDA.AsignarParametroDouble("@PNU_RX_NIV_UMB_NODO_A", PTP.Detalles[0].NivelUmbralNodoA, true);
                baseDatosDA.AsignarParametroDouble("@PNU_SEN_REC_NODO_A", PTP.Detalles[0].SenalRecepcionNodoA, true);
                baseDatosDA.AsignarParametroDouble("@PNU_SEN_REC_NODO_B", PTP.Detalles[0].SenalRecepcionNodoB, true);
                baseDatosDA.AsignarParametroDouble("@PNU_MAR_EFE_DES_NODO_A", PTP.Detalles[0].MargenEfectividadDesvanecimientoNodoA, true);
                baseDatosDA.AsignarParametroDouble("@PNU_MAR_EFE_DES_NODO_B", PTP.Detalles[0].MargenEfectividadDesvanecimientoNodoB, true);
                baseDatosDA.AsignarParametroDouble("@PNU_DIS_ANU_MUL_NODO_A", PTP.Detalles[0].DisponibilidadAnualMultirutasNodoA, true);
                baseDatosDA.AsignarParametroDouble("@PNU_DIS_ANU_LLU_NODO_A", PTP.Detalles[0].DisponibilidadAnualLluviaNodoA, true);
                baseDatosDA.AsignarParametroDouble("@PNU_DIS_ANU_MUL_LLU", PTP.Detalles[0].DisponibilidadAnualMultirutasLluvia, true);
                baseDatosDA.AsignarParametroCadena("@PVC_ID_USUARIO_CRE", PTP.UsuarioCreacion.IdUsuario, true);
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

        public static void EliminarFisicoPTPProceso(PTPBE PTP, DBBaseDatos BaseDatos = null)
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
                baseDatosDA.CrearComando("USP_PTP_PROC", CommandType.StoredProcedure);
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
