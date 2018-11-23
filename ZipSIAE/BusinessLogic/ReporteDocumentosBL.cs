using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Data.SqlClient;
using System.Data;
using System.IO;
using System.Windows.Forms;
using DataAccess;
using System.Globalization;
using BusinessEntity;


namespace BusinessLogic
{

    
    public class ReporteDocumentosBL
    {
       
        

        public void PruebaInterferencia(String IdNodo, String IdTarea, String valorCadena1, String rutaPlantilla)
        {
            DBBaseDatos baseDatosDA = new DBBaseDatos();
            baseDatosDA.Configurar();
            baseDatosDA.Conectar();
            DataTable dt = new DataTable();
            try
            {
                baseDatosDA.CrearComando("USP_R_PRUEBAINTERFERENCIA", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                 dt = baseDatosDA.EjecutarConsultaDataTable();  
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
                String codNodo = dt.Rows[0]["COD_NODO"].ToString();
                byte[] imageBuffer = (byte[])dt.Rows[0]["CAP_PANT_EPMP1000"];
                MemoryStream EPMP1000 = new MemoryStream(imageBuffer);

                String excelGenerado = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";
                File.Copy(rutaPlantilla, excelGenerado, true);

                #region Agregando Valores
                ExcelToolsBL.UpdateCell(excelGenerado, "Selección de Frecuencia", codNodo, 12, "E");
                ExcelToolsBL.AddImageDocument(false, excelGenerado, "Selección de Frecuencia", EPMP1000, "", 18, 3, 696, 394);
                #endregion

                #region Ruta para el Zip
                String rutaNodo = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo;
            String rutaCarpeta = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\ACCESO\\PMP";
                if (Directory.Exists(rutaNodo))
                {
                if (!Directory.Exists(rutaCarpeta))
                {
                    Directory.CreateDirectory(rutaCarpeta);
                    String rutaAlterna = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\ACCESO\\PMP\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";
                    File.Copy(excelGenerado, rutaAlterna, true);
                }
                else
                {
                    String rutaAlterna = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\ACCESO\\PMP\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";
                    File.Copy(excelGenerado, rutaAlterna, true);
                }
                }
                #endregion
          
        }

        public void PruebaServicioDITGPMP(String IdNodo, String IdTarea, String valorCadena1, String rutaPlantilla)
        {
            DBBaseDatos baseDatosDA = new DBBaseDatos();

            baseDatosDA.Configurar();
            baseDatosDA.Conectar();
            DataTable dt = new DataTable();

            try
            {
                baseDatosDA.CrearComando("USP_R_PRUEBAS_SERVICIO_DITG_PMP", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                dt = baseDatosDA.EjecutarConsultaDataTable();


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


            #region Valores
            String TIEMPO_PRUEBA = dt.Rows[0]["TIEMPO_PRUEBA"].ToString();
            String RETARDO_MIN = dt.Rows[0]["RETARDO_MIN"].ToString();
            String RETARDO_MAXIMO = dt.Rows[0]["RETARDO_MAXIMO"].ToString();
            String RETARDO_PROMEDIO = dt.Rows[0]["RETARDO_PROMEDIO"].ToString();
            String JILTER_PROMEDIO = dt.Rows[0]["JILTER_PROMEDIO"].ToString();
            String DESV_ESTANDAR_RETARDO = dt.Rows[0]["DESV_ESTANDAR_RETARDO"].ToString();
            String BYTES_RECIBIDOS = dt.Rows[0]["BYTES_RECIBIDOS"].ToString();
            String THROUGHPUT_PROM = dt.Rows[0]["THROUGHPUT_PROM"].ToString();
            String DESCARTE_PAQUETES = dt.Rows[0]["DESCARTE_PAQUETES"].ToString();

            byte[] FECHA_HORA_ROUTER = (byte[])dt.Rows[0]["FECHA_HORA_ROUTER"];
            MemoryStream FECHA_HORA_ROUTERm = new MemoryStream(FECHA_HORA_ROUTER);
            byte[] DIRECCIONES_MAC = (byte[])dt.Rows[0]["DIRECCIONES_MAC"];
            MemoryStream DIRECCIONES_MACm = new MemoryStream(DIRECCIONES_MAC);
            byte[] RESULTADO_PRUEBA_DITG = (byte[])dt.Rows[0]["RESULTADO_PRUEBA_DITG"];
            MemoryStream RESULTADO_PRUEBA_DITGm = new MemoryStream(RESULTADO_PRUEBA_DITG);
            byte[] PING_CPE_DESDE_NODO_A = (byte[])dt.Rows[0]["PING_CPE_DESDE_NODO_A"];
            MemoryStream PING_CPE_DESDE_NODO_Am = new MemoryStream(PING_CPE_DESDE_NODO_A);
            byte[] PING_ALL_USERS_01 = (byte[])dt.Rows[0]["PING_ALL_USERS_01"];
            MemoryStream PING_ALL_USERS_01m = new MemoryStream(PING_ALL_USERS_01);
            byte[] PING_ALL_USERS_02 = (byte[])dt.Rows[0]["PING_ALL_USERS_02"];
            MemoryStream PING_ALL_USERS_02m = new MemoryStream(PING_ALL_USERS_02);
            byte[] PING_ALL_USERS_03 = (byte[])dt.Rows[0]["PING_ALL_USERS_03"];
            MemoryStream PING_ALL_USERS_03m = new MemoryStream(PING_ALL_USERS_03);
            byte[] PING_ALL_USERS_04 = (byte[])dt.Rows[0]["PING_ALL_USERS_04"];
            MemoryStream PING_ALL_USERS_04m = new MemoryStream(PING_ALL_USERS_04);
            byte[] PING_ALL_USERS_05 = (byte[])dt.Rows[0]["PING_ALL_USERS_05"];
            MemoryStream PING_ALL_USERS_05m = new MemoryStream(PING_ALL_USERS_05);
            byte[] PING_ALL_USERS_06 = (byte[])dt.Rows[0]["PING_ALL_USERS_06"];
            MemoryStream PING_ALL_USERS_06m = new MemoryStream(PING_ALL_USERS_06);

            #endregion

            // String usuarioWindows = Environment.UserName;
            String excelGenerado = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";

            File.Copy(rutaPlantilla, excelGenerado, true);

            #region Agregando valores por Hoja de Excel
            ExcelToolsBL.UpdateCell(excelGenerado, "Pruebas de Servicios", TIEMPO_PRUEBA, 34, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "Pruebas de Servicios", RETARDO_MIN, 35, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "Pruebas de Servicios", RETARDO_MAXIMO, 36, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "Pruebas de Servicios", RETARDO_PROMEDIO, 37, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "Pruebas de Servicios", JILTER_PROMEDIO, 38, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "Pruebas de Servicios", DESV_ESTANDAR_RETARDO, 39, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "Pruebas de Servicios", BYTES_RECIBIDOS, 40, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "Pruebas de Servicios", THROUGHPUT_PROM, 41, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "Pruebas de Servicios", DESCARTE_PAQUETES, 42, "E");

            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Pruebas de Servicios", FECHA_HORA_ROUTERm, "", 72, 2, 720,319);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Pruebas de Servicios", DIRECCIONES_MACm, "", 126, 2, 720, 455);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Pruebas de Servicios", RESULTADO_PRUEBA_DITGm, "", 149, 2, 722, 319);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Pruebas de Servicios", PING_CPE_DESDE_NODO_Am, "", 48, 2, 720, 321);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Pruebas de Servicios", PING_ALL_USERS_01m, "", 98, 2, 232, 191);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Pruebas de Servicios", PING_ALL_USERS_02m, "", 98, 5, 226, 191);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Pruebas de Servicios", PING_ALL_USERS_03m, "", 98, 8, 263, 192);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Pruebas de Servicios", PING_ALL_USERS_04m, "", 110, 2, 232, 184);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Pruebas de Servicios", PING_ALL_USERS_05m, "", 110, 5, 227, 182);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Pruebas de Servicios", PING_ALL_USERS_06m, "", 110, 8, 264, 182);

            #endregion


            #region Ruta para el Zip 
            String rutaNodo = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo;
            if (Directory.Exists(rutaNodo))
            {
                String rutaParcial = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\IE01\\";
                if (!Directory.Exists(rutaParcial)) Directory.CreateDirectory(rutaParcial);
                String rutaAlterna = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\IE01\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";
                File.Copy(excelGenerado, rutaAlterna, true);
            }
          
            #endregion
        }

        public void PruebaServicioDITGPTP(String IdNodo, String IdTarea, String valorCadena1, String rutaPlantilla)
        {
            DBBaseDatos baseDatosDA = new DBBaseDatos();

            baseDatosDA.Configurar();
            baseDatosDA.Conectar();
            DataTable dt = new DataTable();

            try
            {
                baseDatosDA.CrearComando("USP_R_PRUEBAS_SERVICIO_DITG_PTP", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                dt = baseDatosDA.EjecutarConsultaDataTable();

            }
            catch (Exception)
            {

                throw;
            }
            finally
            {

                baseDatosDA.Desconectar();
                baseDatosDA = null;
            }


            #region Valores
            String TIEMPO_PRUEBA = dt.Rows[0]["TIEMPO_PRUEBA"].ToString();
            String RETARDO_MIN = dt.Rows[0]["RETARDO_MIN"].ToString();
            String RETARDO_MAXIMO = dt.Rows[0]["RETARDO_MAXIMO"].ToString();
            String RETARDO_PROMEDIO = dt.Rows[0]["RETARDO_PROMEDIO"].ToString();
            String JILTER_PROMEDIO = dt.Rows[0]["JILTER_PROMEDIO"].ToString();
            String DESV_ESTANDAR_RETARDO = dt.Rows[0]["DESV_ESTANDAR_RETARDO"].ToString();
            String BYTES_RECIBIDOS = dt.Rows[0]["BYTES_RECIBIDOS"].ToString();
            String THROUGHPUT_PROM = dt.Rows[0]["THROUGHPUT_PROM"].ToString();
            String DESCARTE_PAQUETES = dt.Rows[0]["DESCARTE_PAQUETES"].ToString();


            byte[] PANTALLA_RESULT_PRUEBA_DITG = (byte[])dt.Rows[0]["PANTALLA_RESULT_PRUEBA_DITG"];
            MemoryStream mPANTALLA_RESULT_PRUEBA_DITG = new MemoryStream(PANTALLA_RESULT_PRUEBA_DITG);
            byte[] FECHA_HORA_ROUTER = (byte[])dt.Rows[0]["FECHA_HORA_ROUTER"];
            MemoryStream FECHA_HORA_ROUTERm = new MemoryStream(FECHA_HORA_ROUTER);
            byte[] PING_TODOS_USUARIOS_MICRO = (byte[])dt.Rows[0]["PING_TODOS_USUARIOS_MICRO"];
            MemoryStream mPING_TODOS_USUARIOS_MICRO = new MemoryStream(PING_TODOS_USUARIOS_MICRO);
            byte[] DIRECCIONES_MAC = (byte[])dt.Rows[0]["DIRECCIONES_MAC"];
            MemoryStream DIRECCIONES_MACm = new MemoryStream(DIRECCIONES_MAC);
            byte[] RESULTADO_PRUEBA_DITG = (byte[])dt.Rows[0]["RESULTADO_PRUEBA_DITG"];
            MemoryStream RESULTADO_PRUEBA_DITGm = new MemoryStream(RESULTADO_PRUEBA_DITG);

            #endregion

            // String usuarioWindows = Environment.UserName;
            String excelGenerado = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";

            File.Copy(rutaPlantilla, excelGenerado, true);

            #region Agregando valores por Hoja de Excel
            ExcelToolsBL.UpdateCell(excelGenerado, "Pruebas de Servicios", TIEMPO_PRUEBA, 34, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "Pruebas de Servicios", RETARDO_MIN, 35, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "Pruebas de Servicios", RETARDO_MAXIMO, 36, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "Pruebas de Servicios", RETARDO_PROMEDIO, 37, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "Pruebas de Servicios", JILTER_PROMEDIO, 38, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "Pruebas de Servicios", DESV_ESTANDAR_RETARDO, 39, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "Pruebas de Servicios", BYTES_RECIBIDOS, 40, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "Pruebas de Servicios", THROUGHPUT_PROM, 41, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "Pruebas de Servicios", DESCARTE_PAQUETES, 42, "E");

            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Pruebas de Servicios", mPANTALLA_RESULT_PRUEBA_DITG, "", 48,2, 815, 324);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Pruebas de Servicios", FECHA_HORA_ROUTERm, "", 72,2, 815, 324);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Pruebas de Servicios", mPING_TODOS_USUARIOS_MICRO, "", 98,2, 815,379);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Pruebas de Servicios", DIRECCIONES_MACm, "", 126,2,815,462);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Pruebas de Servicios", RESULTADO_PRUEBA_DITGm, "", 151, 2,815, 461);

            #endregion

            #region Ruta para el Zip
            String rutaNodo = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo;
            if (Directory.Exists(rutaNodo))
            {
              try
                {
                    String rutaParcial = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\ACCESO\\PTP_HC-0XX1 (450i)\\";
                    if (!Directory.Exists(rutaParcial)) Directory.CreateDirectory(rutaParcial);
                    String rutaAlterna = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\ACCESO\\PTP_HC-0XX1 (450i)\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";
                    File.Copy(excelGenerado, rutaAlterna, true);
                }
                catch (Exception)
                {
                    String rutaParcial = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\ACCESO\\PTP_HC-0XX2 (alfo)\\";
                    if (!Directory.Exists(rutaParcial)) Directory.CreateDirectory(rutaParcial);
                    String rutaAlterna = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\ACCESO\\PTP_HC-0XX2 (alfo)\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";
                    File.Copy(excelGenerado, rutaAlterna, true);
                }
            }
            #endregion
        }
        public void Anexo2InventarioPMP(String IdNodo, String IdTarea, String valorCadena1, String rutaPlantilla)
        {
            DBBaseDatos baseDatosDA = new DBBaseDatos();
            baseDatosDA.Configurar();
            baseDatosDA.Conectar();
            DataTable dt = new DataTable();
            try
            {
                baseDatosDA.CrearComando("USP_R_ANEXO_02_INVENTARIO_PMP", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                 dt = baseDatosDA.EjecutarConsultaDataTable();          
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                baseDatosDA.Desconectar();
                baseDatosDA = null;
            }
            #region Valores
            String COD_NODO = dt.Rows[0]["COD_NODO"].ToString();
            byte[] ANTENA = (byte[])dt.Rows[0]["ANTENA"];
            MemoryStream ANTENAm = new MemoryStream(ANTENA);
            byte[] ARRESTOR_LAN = (byte[])dt.Rows[0]["ARRESTOR_LAN"];
            MemoryStream ARRESTOR_LANm = new MemoryStream(ARRESTOR_LAN);
            byte[] ODUs = (byte[])dt.Rows[0]["ODUs"];
            MemoryStream ODUsm = new MemoryStream(ODUs);
            byte[] POE = (byte[])dt.Rows[0]["POE"];
            MemoryStream POEm = new MemoryStream(POE);
            #endregion

            String excelGenerado = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";

            File.Copy(rutaPlantilla, excelGenerado, true);

            #region Ingresando Valores
            ExcelToolsBL.UpdateCell(excelGenerado, "11 Serie logística", "NODO: " + COD_NODO, 11, "B");
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "11 Serie logística", ANTENAm, "", 14, 2, 271, 228);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "11 Serie logística", ARRESTOR_LANm, "", 22, 2, 338, 254);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "11 Serie logística", ODUsm, "", 29, 3, 178, 237);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "11 Serie logística", POEm, "", 36, 2, 360, 269);
            #endregion

            #region Ruta para el Zip
            String rutaNodo = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo;
            if (Directory.Exists(rutaNodo))
            {
                String rutaParcial = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\ACCESO\\PMP\\";
                if (!Directory.Exists(rutaParcial)) Directory.CreateDirectory(rutaParcial);
                String rutaAlterna = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\ACCESO\\PMP\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";
                File.Copy(excelGenerado, rutaAlterna, true);
            }
            #endregion
        }

        public void InstalacionPozoTierraTipoA(String IdNodo, String IdTarea, String valorCadena1, String rutaPlantilla)
        {
            DBBaseDatos baseDatosDA = new DBBaseDatos();

            baseDatosDA.Configurar();
            baseDatosDA.Conectar();
            DataTable dt = new DataTable();

            try
            {
                baseDatosDA.CrearComando("USP_R_INSTALACION_POZO_TIERRA_A", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                 dt = baseDatosDA.EjecutarConsultaDataTable();
             
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

            #region valores String
            String TIPO_INSTITUCION = dt.Rows[0]["TIPO_INSTITUCION"].ToString();
            String CODIGO_IIBB = dt.Rows[0]["CODIGO_IIBB"].ToString();
            String NOMBRE_IIBB = dt.Rows[0]["NOMBRE_IIBB"].ToString();
            #endregion

            #region valores binarios

            #region Campos que no se usan en este tipo de pozo
            //byte[] CINCO_OHM_FOTOGRAFIA_FRONTAL_iIBB = (byte[])ds.Tables[0].Rows[0]["CINCO_OHM_FOTOGRAFIA_FRONTAL_iIBB"];
            //MemoryStream CINCO_OHM_FOTOGRAFIA_FRONTAL_iIBBm = new MemoryStream(CINCO_OHM_FOTOGRAFIA_FRONTAL_iIBB);
            //byte[] CINCO_OHM_UBICACION_POZO_ANTES_INSTALAR = (byte[])ds.Tables[0].Rows[0]["CINCO_OHM_UBICACION_POZO_ANTES_INSTALAR"];
            //MemoryStream CINCO_OHM_UBICACION_POZO_ANTES_INSTALARm = new MemoryStream(CINCO_OHM_UBICACION_POZO_ANTES_INSTALAR);
            //byte[] CINCO_OHM_PAN_ZANJA_ABIERTA = (byte[])ds.Tables[0].Rows[0]["CINCO_OHM_PAN_ZANJA_ABIERTA"];
            //MemoryStream CINCO_OHM_PAN_ZANJA_ABIERTAm = new MemoryStream(CINCO_OHM_PAN_ZANJA_ABIERTA);
            //byte[] CINCO_OHM_PAN_VERTIDO_TIERRA = (byte[])ds.Tables[0].Rows[0]["CINCO_OHM_PAN_VERTIDO_TIERRA"];
            //MemoryStream CINCO_OHM_PAN_VERTIDO_TIERRAm = new MemoryStream(CINCO_OHM_PAN_VERTIDO_TIERRA);
            //byte[] CINCO_OHM_PAN_VERTIDO_SAL = (byte[])ds.Tables[0].Rows[0]["CINCO_OHM_PAN_VERTIDO_SAL"];
            //MemoryStream CINCO_OHM_PAN_VERTIDO_SALm = new MemoryStream(CINCO_OHM_PAN_VERTIDO_SAL);
            //byte[] CINCO_OHM_PAN_VERTIDO_DISOLUCION = (byte[])ds.Tables[0].Rows[0]["CINCO_OHM_PAN_VERTIDO_DISOLUCION"];
            //MemoryStream CINCO_OHM_PAN_VERTIDO_DISOLUCIONm = new MemoryStream(CINCO_OHM_PAN_VERTIDO_DISOLUCION);
            //byte[] CINCO_OHM_PAN_COL_REJE_COBRE01 = (byte[])ds.Tables[0].Rows[0]["CINCO_OHM_PAN_COL_REJE_COBRE01"];
            //MemoryStream CINCO_OHM_PAN_COL_REJE_COBRE01m = new MemoryStream(CINCO_OHM_PAN_COL_REJE_COBRE01);
            //byte[] CINCO_OHM_PAN_COL_REJE_COBRE02 = (byte[])ds.Tables[0].Rows[0]["CINCO_OHM_PAN_COL_REJE_COBRE02"];
            //MemoryStream CINCO_OHM_PAN_COL_REJE_COBRE02m = new MemoryStream(CINCO_OHM_PAN_COL_REJE_COBRE02);
            //byte[] CINCO_OHM_PAN_VERTIDO_DISOLUCION_SOBRE_REJE = (byte[])ds.Tables[0].Rows[0]["CINCO_OHM_PAN_VERTIDO_DISOLUCION_SOBRE_REJE"];
            //MemoryStream CINCO_OHM_PAN_VERTIDO_DISOLUCION_SOBRE_REJEm = new MemoryStream(CINCO_OHM_PAN_VERTIDO_DISOLUCION_SOBRE_REJE);
            //byte[] CINCO_OHM_PAN_VERTIDO_SAL_GRANULADO_LUEGO_DEL_CEMENTO = (byte[])ds.Tables[0].Rows[0]["CINCO_OHM_PAN_VERTIDO_SAL_GRANULADO_LUEGO_DEL_CEMENTO"];
            //MemoryStream CINCO_OHM_PAN_VERTIDO_SAL_GRANULADO_LUEGO_DEL_CEMENTOm = new MemoryStream(CINCO_OHM_PAN_VERTIDO_SAL_GRANULADO_LUEGO_DEL_CEMENTO);
            //byte[] CINCO_OHM_PAN_VERTIDO_RESTO_TIERRA_CULTIVO = (byte[])ds.Tables[0].Rows[0]["CINCO_OHM_PAN_VERTIDO_RESTO_TIERRA_CULTIVO"];
            //MemoryStream CINCO_OHM_PAN_VERTIDO_RESTO_TIERRA_CULTIVOm = new MemoryStream(CINCO_OHM_PAN_VERTIDO_RESTO_TIERRA_CULTIVO);
            //byte[] CINCO_OHM_PAN_RELLENADO_TIERRA_CERNIDA_ZANJA = (byte[])ds.Tables[0].Rows[0]["CINCO_OHM_PAN_RELLENADO_TIERRA_CERNIDA_ZANJA"];
            //MemoryStream CINCO_OHM_PAN_RELLENADO_TIERRA_CERNIDA_ZANJAm = new MemoryStream(CINCO_OHM_PAN_RELLENADO_TIERRA_CERNIDA_ZANJA);
            //byte[] CINCO_OHM_MED1_PAN_POZO_TIERRA = (byte[])ds.Tables[0].Rows[0]["CINCO_OHM_MED1_PAN_POZO_TIERRA"];
            //MemoryStream CINCO_OHM_MED1_PAN_POZO_TIERRAm = new MemoryStream(CINCO_OHM_MED1_PAN_POZO_TIERRA);
            //byte[] CINCO_OHM_MED2_PAN_POZO_TIERRA = (byte[])ds.Tables[0].Rows[0]["CINCO_OHM_MED2_PAN_POZO_TIERRA"];
            //MemoryStream CINCO_OHM_MED2_PAN_POZO_TIERRAm = new MemoryStream(CINCO_OHM_MED2_PAN_POZO_TIERRA);
            //byte[] CINCO_OHM_MED3_PAN_POZO_TIERRA = (byte[])ds.Tables[0].Rows[0]["CINCO_OHM_MED3_PAN_POZO_TIERRA"];
            //MemoryStream CINCO_OHM_MED3_PAN_POZO_TIERRAm = new MemoryStream(CINCO_OHM_MED3_PAN_POZO_TIERRA);
            #endregion

            byte[] DIEZ_OHM_FRONTAL_IIBB = (byte[])dt.Rows[0]["DIEZ_OHM_FRONTAL_IIBB"];
            MemoryStream DIEZ_OHM_FRONTAL_IIBBm = new MemoryStream(DIEZ_OHM_FRONTAL_IIBB);
            byte[] DIEZ_OHM_UBIC_POZO_ANTES_INSTALACION = (byte[])dt.Rows[0]["DIEZ_OHM_UBIC_POZO_ANTES_INSTALACION"];
            MemoryStream DIEZ_OHM_UBIC_POZO_ANTES_INSTALACIONm = new MemoryStream(DIEZ_OHM_UBIC_POZO_ANTES_INSTALACION);
            byte[] DIEZ_OHM_PAN_ZANJA_ABIERTA = (byte[])dt.Rows[0]["DIEZ_OHM_PAN_ZANJA_ABIERTA"];
            MemoryStream DIEZ_OHM_PAN_ZANJA_ABIERTAm = new MemoryStream(DIEZ_OHM_PAN_ZANJA_ABIERTA);
            byte[] DIEZ_OHM_PAN_VERITDO_TIERRA_EN_ZANJA = (byte[])dt.Rows[0]["DIEZ_OHM_PAN_VERITDO_TIERRA_EN_ZANJA"];
            MemoryStream DIEZ_OHM_PAN_VERITDO_TIERRA_EN_ZANJAm = new MemoryStream(DIEZ_OHM_PAN_VERITDO_TIERRA_EN_ZANJA);
            byte[] DIEZ_OHM_PAN_VERTIDO_SAL_GRANULADA_ZANJA = (byte[])dt.Rows[0]["DIEZ_OHM_PAN_VERTIDO_SAL_GRANULADA_ZANJA"];
            MemoryStream DIEZ_OHM_PAN_VERTIDO_SAL_GRANULADA_ZANJAm = new MemoryStream(DIEZ_OHM_PAN_VERTIDO_SAL_GRANULADA_ZANJA);
            byte[] DIEZ_OHM_VESTIDO_DIS_CEMENTO = (byte[])dt.Rows[0]["DIEZ_OHM_VESTIDO_DIS_CEMENTO"];
            MemoryStream DIEZ_OHM_VESTIDO_DIS_CEMENTOm = new MemoryStream(DIEZ_OHM_VESTIDO_DIS_CEMENTO);
            byte[] DIEZ_OHM_PAN_REJE_COBRE_01 = (byte[])dt.Rows[0]["DIEZ_OHM_PAN_REJE_COBRE_01"];
            MemoryStream DIEZ_OHM_PAN_REJE_COBRE_01m = new MemoryStream(DIEZ_OHM_PAN_REJE_COBRE_01);
            byte[] DIEZ_OHM_PAN_REJE_COBRE_02 = (byte[])dt.Rows[0]["DIEZ_OHM_PAN_REJE_COBRE_02"];
            MemoryStream DIEZ_OHM_PAN_REJE_COBRE_02m = new MemoryStream(DIEZ_OHM_PAN_REJE_COBRE_02);
            byte[] DIEZ_OHM_PAN_VERTIDO_DIS_ZANJA = (byte[])dt.Rows[0]["DIEZ_OHM_PAN_VERTIDO_DIS_ZANJA"];
            MemoryStream DIEZ_OHM_PAN_VERTIDO_DIS_ZANJAm = new MemoryStream(DIEZ_OHM_PAN_VERTIDO_DIS_ZANJA);
            byte[] DIEZ_OHM_PAN_VERTIDO_sAL_GRANULADA = (byte[])dt.Rows[0]["DIEZ_OHM_PAN_VERTIDO_sAL_GRANULADA"];
            MemoryStream DIEZ_OHM_PAN_VERTIDO_sAL_GRANULADAm = new MemoryStream(DIEZ_OHM_PAN_VERTIDO_sAL_GRANULADA);
            byte[] DIEZ_OHM_VERTIDO_RESTO_TIERRA = (byte[])dt.Rows[0]["DIEZ_OHM_VERTIDO_RESTO_TIERRA"];
            MemoryStream DIEZ_OHM_VERTIDO_RESTO_TIERRAm = new MemoryStream(DIEZ_OHM_VERTIDO_RESTO_TIERRA);
            byte[] DIEZ_OHM_VERTIDO_RELLENADO_TIERRA = (byte[])dt.Rows[0]["DIEZ_OHM_VERTIDO_RELLENADO_TIERRA"];
            MemoryStream DIEZ_OHM_VERTIDO_RELLENADO_TIERRAm = new MemoryStream(DIEZ_OHM_VERTIDO_RELLENADO_TIERRA);
            byte[] DIEZ_OHM_MEDICION1 = (byte[])dt.Rows[0]["DIEZ_OHM_MEDICION1"];
            MemoryStream DIEZ_OHM_MEDICION1m = new MemoryStream(DIEZ_OHM_MEDICION1);
            byte[] DIEZ_OHM_MEDICION2 = (byte[])dt.Rows[0]["DIEZ_OHM_MEDICION2"];
            MemoryStream DIEZ_OHM_MEDICION2m = new MemoryStream(DIEZ_OHM_MEDICION2);
            byte[] DIEZ_OHM_MEDICION3 = (byte[])dt.Rows[0]["DIEZ_OHM_MEDICION3"];
            MemoryStream DIEZ_OHM_MEDICION3m = new MemoryStream(DIEZ_OHM_MEDICION3);
            #endregion

            // String usuarioWindows = Environment.UserName;
            String excelGenerado = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";
            File.Copy(rutaPlantilla, excelGenerado, true);

            #region Agregando los datos 
            ExcelToolsBL.UpdateCell(excelGenerado, "POZO A TIERRA 10 Ohm", TIPO_INSTITUCION, 7, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "POZO A TIERRA 10 Ohm", CODIGO_IIBB, 7, "N");
            ExcelToolsBL.UpdateCell(excelGenerado, "POZO A TIERRA 10 Ohm", NOMBRE_IIBB, 8, "G");

            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 10 Ohm", DIEZ_OHM_FRONTAL_IIBBm, "", 13, 3, 1834, 465);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 10 Ohm", DIEZ_OHM_UBIC_POZO_ANTES_INSTALACIONm, "", 49, 3, 1835, 410);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 10 Ohm", DIEZ_OHM_PAN_ZANJA_ABIERTAm, "", 83, 3, 723, 452);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 10 Ohm", DIEZ_OHM_PAN_VERITDO_TIERRA_EN_ZANJAm, "", 83, 14, 913, 447);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 10 Ohm", DIEZ_OHM_PAN_VERTIDO_SAL_GRANULADA_ZANJAm, "", 102, 3, 720, 450);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 10 Ohm", DIEZ_OHM_VESTIDO_DIS_CEMENTOm, "", 102, 14, 912, 451);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 10 Ohm", DIEZ_OHM_PAN_REJE_COBRE_01m, "", 123, 3, 317, 452);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 10 Ohm", DIEZ_OHM_PAN_REJE_COBRE_02m, "", 123, 8, 402, 447);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 10 Ohm", DIEZ_OHM_VESTIDO_DIS_CEMENTOm, "", 123, 14, 914, 450);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 10 Ohm", DIEZ_OHM_PAN_VERTIDO_sAL_GRANULADAm, "", 142, 3, 721, 447);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 10 Ohm", DIEZ_OHM_VERTIDO_RESTO_TIERRAm, "", 142, 14, 912, 450);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 10 Ohm", DIEZ_OHM_VERTIDO_RELLENADO_TIERRAm, "", 160, 3, 727, 447);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 10 Ohm", DIEZ_OHM_MEDICION1m, "", 181, 3, 1835, 360);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 10 Ohm", DIEZ_OHM_MEDICION2m, "", 201, 3, 1838, 366);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 10 Ohm", DIEZ_OHM_MEDICION3m, "", 221, 3, 1837, 372);
            #endregion

            #region Ruta para el Zip
            String rutaNodo = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo;
            if (Directory.Exists(rutaNodo))
            {
                String rutaParcial = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\CS01\\";
                if (!Directory.Exists(rutaParcial)) Directory.CreateDirectory(rutaParcial);
                String rutaAlterna = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\CS01\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";
                File.Copy(excelGenerado, rutaAlterna, true);
            }
            #endregion
        }

        public void InstalacionPozoTierraTipoB(String IdNodo, String IdTarea, String valorCadena1, String rutaPlantilla)
        {
            DBBaseDatos baseDatosDA = new DBBaseDatos();

            baseDatosDA.Configurar();
            baseDatosDA.Conectar();
            DataTable dt = new DataTable();

            try
            {
                baseDatosDA.CrearComando("USP_R_INSTALACION_POZO_TIERRA_B", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                 dt = baseDatosDA.EjecutarConsultaDataTable();         
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
            #region valores String
            String TIPO_INSTITUCION = dt.Rows[0]["TIPO_INSTITUCION"].ToString();
            String CODIGO_IIBB = dt.Rows[0]["CODIGO_IIBB"].ToString();
            String NOMBRE_IIBB = dt.Rows[0]["NOMBRE_IIBB"].ToString();
            #endregion

            #region valores binarios
            byte[] CINCO_OHM_FOTOGRAFIA_FRONTAL_iIBB = (byte[])dt.Rows[0]["CINCO_OHM_FOTOGRAFIA_FRONTAL_iIBB"];
            MemoryStream CINCO_OHM_FOTOGRAFIA_FRONTAL_iIBBm = new MemoryStream(CINCO_OHM_FOTOGRAFIA_FRONTAL_iIBB);
            byte[] CINCO_OHM_UBICACION_POZO_ANTES_INSTALAR = (byte[])dt.Rows[0]["CINCO_OHM_UBICACION_POZO_ANTES_INSTALAR"];
            MemoryStream CINCO_OHM_UBICACION_POZO_ANTES_INSTALARm = new MemoryStream(CINCO_OHM_UBICACION_POZO_ANTES_INSTALAR);
            byte[] CINCO_OHM_PAN_ZANJA_ABIERTA = (byte[])dt.Rows[0]["CINCO_OHM_PAN_ZANJA_ABIERTA"];
            MemoryStream CINCO_OHM_PAN_ZANJA_ABIERTAm = new MemoryStream(CINCO_OHM_PAN_ZANJA_ABIERTA);
            byte[] CINCO_OHM_PAN_VERTIDO_TIERRA = (byte[])dt.Rows[0]["CINCO_OHM_PAN_VERTIDO_TIERRA"];
            MemoryStream CINCO_OHM_PAN_VERTIDO_TIERRAm = new MemoryStream(CINCO_OHM_PAN_VERTIDO_TIERRA);
            byte[] CINCO_OHM_PAN_VERTIDO_SAL = (byte[])dt.Rows[0]["CINCO_OHM_PAN_VERTIDO_SAL"];
            MemoryStream CINCO_OHM_PAN_VERTIDO_SALm = new MemoryStream(CINCO_OHM_PAN_VERTIDO_SAL);
            byte[] CINCO_OHM_PAN_VERTIDO_DISOLUCION = (byte[])dt.Rows[0]["CINCO_OHM_PAN_VERTIDO_DISOLUCION"];
            MemoryStream CINCO_OHM_PAN_VERTIDO_DISOLUCIONm = new MemoryStream(CINCO_OHM_PAN_VERTIDO_DISOLUCION);
            byte[] CINCO_OHM_PAN_COL_REJE_COBRE01 = (byte[])dt.Rows[0]["CINCO_OHM_PAN_COL_REJE_COBRE01"];
            MemoryStream CINCO_OHM_PAN_COL_REJE_COBRE01m = new MemoryStream(CINCO_OHM_PAN_COL_REJE_COBRE01);
            byte[] CINCO_OHM_PAN_COL_REJE_COBRE02 = (byte[])dt.Rows[0]["CINCO_OHM_PAN_COL_REJE_COBRE02"];
            MemoryStream CINCO_OHM_PAN_COL_REJE_COBRE02m = new MemoryStream(CINCO_OHM_PAN_COL_REJE_COBRE02);
            byte[] CINCO_OHM_PAN_VERTIDO_DISOLUCION_SOBRE_REJE = (byte[])dt.Rows[0]["CINCO_OHM_PAN_VERTIDO_DISOLUCION_SOBRE_REJE"];
            MemoryStream CINCO_OHM_PAN_VERTIDO_DISOLUCION_SOBRE_REJEm = new MemoryStream(CINCO_OHM_PAN_VERTIDO_DISOLUCION_SOBRE_REJE);
            byte[] CINCO_OHM_PAN_VERTIDO_SAL_GRANULADO_LUEGO_DEL_CEMENTO = (byte[])dt.Rows[0]["CINCO_OHM_PAN_VERTIDO_SAL_GRANULADO_LUEGO_DEL_CEMENTO"];
            MemoryStream CINCO_OHM_PAN_VERTIDO_SAL_GRANULADO_LUEGO_DEL_CEMENTOm = new MemoryStream(CINCO_OHM_PAN_VERTIDO_SAL_GRANULADO_LUEGO_DEL_CEMENTO);
            byte[] CINCO_OHM_PAN_VERTIDO_RESTO_TIERRA_CULTIVO = (byte[])dt.Rows[0]["CINCO_OHM_PAN_VERTIDO_RESTO_TIERRA_CULTIVO"];
            MemoryStream CINCO_OHM_PAN_VERTIDO_RESTO_TIERRA_CULTIVOm = new MemoryStream(CINCO_OHM_PAN_VERTIDO_RESTO_TIERRA_CULTIVO);
            byte[] CINCO_OHM_PAN_RELLENADO_TIERRA_CERNIDA_ZANJA = (byte[])dt.Rows[0]["CINCO_OHM_PAN_RELLENADO_TIERRA_CERNIDA_ZANJA"];
            MemoryStream CINCO_OHM_PAN_RELLENADO_TIERRA_CERNIDA_ZANJAm = new MemoryStream(CINCO_OHM_PAN_RELLENADO_TIERRA_CERNIDA_ZANJA);
            byte[] CINCO_OHM_MED1_PAN_POZO_TIERRA = (byte[])dt.Rows[0]["CINCO_OHM_MED1_PAN_POZO_TIERRA"];
            MemoryStream CINCO_OHM_MED1_PAN_POZO_TIERRAm = new MemoryStream(CINCO_OHM_MED1_PAN_POZO_TIERRA);
            byte[] CINCO_OHM_MED2_PAN_POZO_TIERRA = (byte[])dt.Rows[0]["CINCO_OHM_MED2_PAN_POZO_TIERRA"];
            MemoryStream CINCO_OHM_MED2_PAN_POZO_TIERRAm = new MemoryStream(CINCO_OHM_MED2_PAN_POZO_TIERRA);
            byte[] CINCO_OHM_MED3_PAN_POZO_TIERRA = (byte[])dt.Rows[0]["CINCO_OHM_MED3_PAN_POZO_TIERRA"];
            MemoryStream CINCO_OHM_MED3_PAN_POZO_TIERRAm = new MemoryStream(CINCO_OHM_MED3_PAN_POZO_TIERRA);
            byte[] DIEZ_OHM_FRONTAL_IIBB = (byte[])dt.Rows[0]["DIEZ_OHM_FRONTAL_IIBB"];
            MemoryStream DIEZ_OHM_FRONTAL_IIBBm = new MemoryStream(DIEZ_OHM_FRONTAL_IIBB);
            byte[] DIEZ_OHM_UBIC_POZO_ANTES_INSTALACION = (byte[])dt.Rows[0]["DIEZ_OHM_UBIC_POZO_ANTES_INSTALACION"];
            MemoryStream DIEZ_OHM_UBIC_POZO_ANTES_INSTALACIONm = new MemoryStream(DIEZ_OHM_UBIC_POZO_ANTES_INSTALACION);
            byte[] DIEZ_OHM_PAN_ZANJA_ABIERTA = (byte[])dt.Rows[0]["DIEZ_OHM_PAN_ZANJA_ABIERTA"];
            MemoryStream DIEZ_OHM_PAN_ZANJA_ABIERTAm = new MemoryStream(DIEZ_OHM_PAN_ZANJA_ABIERTA);
            byte[] DIEZ_OHM_PAN_VERITDO_TIERRA_EN_ZANJA = (byte[])dt.Rows[0]["DIEZ_OHM_PAN_VERITDO_TIERRA_EN_ZANJA"];
            MemoryStream DIEZ_OHM_PAN_VERITDO_TIERRA_EN_ZANJAm = new MemoryStream(DIEZ_OHM_PAN_VERITDO_TIERRA_EN_ZANJA);
            byte[] DIEZ_OHM_PAN_VERTIDO_SAL_GRANULADA_ZANJA = (byte[])dt.Rows[0]["DIEZ_OHM_PAN_VERTIDO_SAL_GRANULADA_ZANJA"];
            MemoryStream DIEZ_OHM_PAN_VERTIDO_SAL_GRANULADA_ZANJAm = new MemoryStream(DIEZ_OHM_PAN_VERTIDO_SAL_GRANULADA_ZANJA);
            byte[] DIEZ_OHM_VESTIDO_DIS_CEMENTO = (byte[])dt.Rows[0]["DIEZ_OHM_VESTIDO_DIS_CEMENTO"];
            MemoryStream DIEZ_OHM_VESTIDO_DIS_CEMENTOm = new MemoryStream(DIEZ_OHM_VESTIDO_DIS_CEMENTO);
            byte[] DIEZ_OHM_PAN_REJE_COBRE_01 = (byte[])dt.Rows[0]["DIEZ_OHM_PAN_REJE_COBRE_01"];
            MemoryStream DIEZ_OHM_PAN_REJE_COBRE_01m = new MemoryStream(DIEZ_OHM_PAN_REJE_COBRE_01);
            byte[] DIEZ_OHM_PAN_REJE_COBRE_02 = (byte[])dt.Rows[0]["DIEZ_OHM_PAN_REJE_COBRE_02"];
            MemoryStream DIEZ_OHM_PAN_REJE_COBRE_02m = new MemoryStream(DIEZ_OHM_PAN_REJE_COBRE_02);
            byte[] DIEZ_OHM_PAN_VERTIDO_DIS_ZANJA = (byte[])dt.Rows[0]["DIEZ_OHM_PAN_VERTIDO_DIS_ZANJA"];
            MemoryStream DIEZ_OHM_PAN_VERTIDO_DIS_ZANJAm = new MemoryStream(DIEZ_OHM_PAN_VERTIDO_DIS_ZANJA);
            byte[] DIEZ_OHM_PAN_VERTIDO_sAL_GRANULADA = (byte[])dt.Rows[0]["DIEZ_OHM_PAN_VERTIDO_sAL_GRANULADA"];
            MemoryStream DIEZ_OHM_PAN_VERTIDO_sAL_GRANULADAm = new MemoryStream(DIEZ_OHM_PAN_VERTIDO_sAL_GRANULADA);
            byte[] DIEZ_OHM_VERTIDO_RESTO_TIERRA = (byte[])dt.Rows[0]["DIEZ_OHM_VERTIDO_RESTO_TIERRA"];
            MemoryStream DIEZ_OHM_VERTIDO_RESTO_TIERRAm = new MemoryStream(DIEZ_OHM_VERTIDO_RESTO_TIERRA);
            byte[] DIEZ_OHM_VERTIDO_RELLENADO_TIERRA = (byte[])dt.Rows[0]["DIEZ_OHM_VERTIDO_RELLENADO_TIERRA"];
            MemoryStream DIEZ_OHM_VERTIDO_RELLENADO_TIERRAm = new MemoryStream(DIEZ_OHM_VERTIDO_RELLENADO_TIERRA);
            byte[] DIEZ_OHM_MEDICION1 = (byte[])dt.Rows[0]["DIEZ_OHM_MEDICION1"];
            MemoryStream DIEZ_OHM_MEDICION1m = new MemoryStream(DIEZ_OHM_MEDICION1);
            byte[] DIEZ_OHM_MEDICION2 = (byte[])dt.Rows[0]["DIEZ_OHM_MEDICION2"];
            MemoryStream DIEZ_OHM_MEDICION2m = new MemoryStream(DIEZ_OHM_MEDICION2);
            byte[] DIEZ_OHM_MEDICION3 = (byte[])dt.Rows[0]["DIEZ_OHM_MEDICION3"];
            MemoryStream DIEZ_OHM_MEDICION3m = new MemoryStream(DIEZ_OHM_MEDICION3);
            #endregion

            // String usuarioWindows = Environment.UserName;
            String excelGenerado = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";
            File.Copy(rutaPlantilla, excelGenerado, true);

            #region Agregando los datos 

            #region Pozo a Tierra 5 Ohm

            ExcelToolsBL.UpdateCell(excelGenerado, "POZO A TIERRA 5 Ohm", TIPO_INSTITUCION, 7, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "POZO A TIERRA 5 Ohm", CODIGO_IIBB, 7, "N");
            ExcelToolsBL.UpdateCell(excelGenerado, "POZO A TIERRA 5 Ohm", NOMBRE_IIBB, 8, "G");

            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 5 Ohm", DIEZ_OHM_FRONTAL_IIBBm, "", 13, 3, 1860, 448);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 5 Ohm", DIEZ_OHM_UBIC_POZO_ANTES_INSTALACIONm, "", 49, 3, 1860, 397);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 5 Ohm", DIEZ_OHM_PAN_ZANJA_ABIERTAm, "", 83, 3,740, 445);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 5 Ohm", DIEZ_OHM_PAN_VERITDO_TIERRA_EN_ZANJAm, "", 83, 14, 916, 445);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 5 Ohm", DIEZ_OHM_PAN_VERTIDO_SAL_GRANULADA_ZANJAm, "", 102, 3, 740, 440);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 5 Ohm", DIEZ_OHM_VESTIDO_DIS_CEMENTOm, "", 102, 14, 918, 440);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 5 Ohm", DIEZ_OHM_PAN_REJE_COBRE_01m, "", 123, 3, 337, 447);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 5 Ohm", DIEZ_OHM_PAN_REJE_COBRE_02m, "", 123, 8, 407, 444);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 5 Ohm", DIEZ_OHM_VESTIDO_DIS_CEMENTOm, "", 123, 14, 916, 444);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 5 Ohm", DIEZ_OHM_PAN_VERTIDO_sAL_GRANULADAm, "", 142, 3, 740, 440);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 5 Ohm", DIEZ_OHM_VERTIDO_RESTO_TIERRAm, "", 142, 14, 916, 440);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 5 Ohm", DIEZ_OHM_VERTIDO_RELLENADO_TIERRAm, "", 160, 3, 740, 440);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 5 Ohm", DIEZ_OHM_MEDICION1m, "", 181, 3, 1860, 351);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 5 Ohm", DIEZ_OHM_MEDICION2m, "", 201, 3, 1859, 351);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 5 Ohm", DIEZ_OHM_MEDICION3m, "", 221, 3, 1859, 351);

            #endregion

            #region Pozo a Tierra 10 ohm

            ExcelToolsBL.UpdateCell(excelGenerado, "POZO A TIERRA 10 Ohm", TIPO_INSTITUCION, 7, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "POZO A TIERRA 10 Ohm", CODIGO_IIBB, 7, "N");
            ExcelToolsBL.UpdateCell(excelGenerado, "POZO A TIERRA 10 Ohm", NOMBRE_IIBB, 8, "G");

            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 10 Ohm", CINCO_OHM_FOTOGRAFIA_FRONTAL_iIBBm, "", 13, 3, 1835, 467);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 10 Ohm", CINCO_OHM_UBICACION_POZO_ANTES_INSTALARm, "", 49, 3, 1835, 413);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 10 Ohm", CINCO_OHM_PAN_ZANJA_ABIERTAm, "", 83, 3, 721, 450);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 10 Ohm", CINCO_OHM_PAN_VERTIDO_TIERRAm, "", 83, 14, 912, 451);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 10 Ohm", CINCO_OHM_PAN_VERTIDO_SALm, "", 102, 3, 720, 448);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 10 Ohm", CINCO_OHM_PAN_VERTIDO_DISOLUCIONm, "", 102, 14, 914, 448);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 10 Ohm", CINCO_OHM_PAN_COL_REJE_COBRE01m, "", 123, 3,314, 450);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 10 Ohm", CINCO_OHM_PAN_COL_REJE_COBRE02m, "", 123, 8, 406, 452);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 10 Ohm", CINCO_OHM_PAN_VERTIDO_DISOLUCION_SOBRE_REJEm, "", 123, 14, 912, 452);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 10 Ohm", CINCO_OHM_PAN_VERTIDO_SAL_GRANULADO_LUEGO_DEL_CEMENTOm, "", 142, 3,720,450);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 10 Ohm", CINCO_OHM_PAN_VERTIDO_RESTO_TIERRA_CULTIVOm, "", 142, 14, 914, 450);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 10 Ohm", CINCO_OHM_PAN_RELLENADO_TIERRA_CERNIDA_ZANJAm, "", 160, 3, 720, 449);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 10 Ohm", CINCO_OHM_MED1_PAN_POZO_TIERRAm, "", 181,3, 1834, 361);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 10 Ohm", CINCO_OHM_MED2_PAN_POZO_TIERRAm, "", 201, 3, 1834, 367);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "POZO A TIERRA 10 Ohm", CINCO_OHM_MED3_PAN_POZO_TIERRAm, "", 221,3, 1835, 372);

            #endregion
            #endregion

            #region Ruta Zip
            String rutaNodo = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo;
            if (Directory.Exists(rutaNodo))
            {
                String rutaParcial = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\IE01\\";
                if (!Directory.Exists(rutaParcial)) Directory.CreateDirectory(rutaParcial);
                String rutaAlterna = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\IE01\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";
                File.Copy(excelGenerado, rutaAlterna, true);
            }
            #endregion
        }

        public void ActaSeguridadDistribucion(String IdNodo, String IdTarea, String valorCadena1, String rutaPlantilla)
        {
            DBBaseDatos baseDatosDA = new DBBaseDatos();

            baseDatosDA.Configurar();
            baseDatosDA.Conectar();
            DataTable dt = new DataTable();
            DataTable dt1 = new DataTable();
            DataTable dt2 = new DataTable();

            try
            {
                baseDatosDA.CrearComando("USP_R_ACTA_SEGURIDAD_DISTRIBUCION", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                dt = baseDatosDA.EjecutarConsultaDataTable();

                baseDatosDA.CrearComando("USP_R_EQUIPAMIENTOS_SEGURIDAD_DISTRIBUCION", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                dt1 = baseDatosDA.EjecutarConsultaDataTable();

                baseDatosDA.CrearComando("USP_R_MATERIALES_SEGURIDAD_DISTRIBUCION", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                dt2 = baseDatosDA.EjecutarConsultaDataTable();


        
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {

                baseDatosDA.Desconectar();
                baseDatosDA = null;
            }
            #region valores_String
            String NOMBRE_NODO = "NODO " + dt.Rows[0]["NOMBRE_NODO"].ToString();
            String CODIGO_NODO = dt.Rows[0]["CODIGO_NODO"].ToString();
            String TIPO_NODO = dt.Rows[0]["TIPO_NODO"].ToString();

            String fechaSQL = dt.Rows[0]["FECHA"].ToString();
            String FECHA = "";
            if (fechaSQL != "")
            {
                DateTime dtFecha = DateTime.Parse(fechaSQL);
                FECHA = dtFecha.ToString("dd/MM/yyyy");
            }
            else { FECHA = ""; }

            String fechaSQL_1 = dt.Rows[0]["EXTINGUIDOR_EXT_FECHA_EXPIRACION"].ToString();
            String EXTINGUIDOR_EXT_FECHA_EXPIRACION = "";
            if (fechaSQL_1 != "")
            {
                DateTime dtFecha = DateTime.Parse(fechaSQL_1);
                EXTINGUIDOR_EXT_FECHA_EXPIRACION = dtFecha.ToString("dd/MM/yyyy");
            }
            else { EXTINGUIDOR_EXT_FECHA_EXPIRACION = ""; }

            String FechaSQL_2 = dt.Rows[0]["EXTINGUIDOR_INT_FECHA_EXPIRACION"].ToString();
            String EXTINGUIDOR_INT_FECHA_EXPIRACION = "";
            if (FechaSQL_2 != "")
            {
                DateTime dtFecha = DateTime.Parse(FechaSQL_2);
                EXTINGUIDOR_INT_FECHA_EXPIRACION = dtFecha.ToString("dd/MM/yyyy");
            }
            else { EXTINGUIDOR_INT_FECHA_EXPIRACION = ""; }
            String DEPARTAMENTO = dt.Rows[0]["DEPARTAMENTO"].ToString();
            String SERIAL_CONTROLADOR = dt.Rows[0]["SERIAL_CONTROLADOR"].ToString();
            //IP CONTROLADOR
            String IP_CONTROLADOR = dt.Rows[0]["IP_CONTROLADOR"].ToString();


            #endregion

            #region valores binarios
            byte[] FACHADA_DEL_NODO = (byte[])dt.Rows[0]["FACHADA_DEL_NODO"];
            MemoryStream mFACHADA_DEL_NODO = new MemoryStream(FACHADA_DEL_NODO);
            byte[] SALA_EQUIPOS_PANORAMICA_RACK = (byte[])dt.Rows[0]["SALA_EQUIPOS_PANORAMICA_RACK"];
            MemoryStream mSALA_EQUIPOS_PANORAMICA_RACK = new MemoryStream(SALA_EQUIPOS_PANORAMICA_RACK);
            byte[] PANORAMICA_INTERIOR_01 = (byte[])dt.Rows[0]["PANORAMICA_INTERIOR_01"];
            MemoryStream mPANORAMICA_INTERIOR_01 = new MemoryStream(PANORAMICA_INTERIOR_01);
            byte[] PANORAMICA_INTERIOR_02 = (byte[])dt.Rows[0]["PANORAMICA_INTERIOR_02"];
            MemoryStream mPANORAMICA_INTERIOR_02 = new MemoryStream(PANORAMICA_INTERIOR_02);
            byte[] PANORAMICA_EQUIPOS_PATIO = (byte[])dt.Rows[0]["PANORAMICA_EQUIPOS_PATIO"];
            MemoryStream mPANORAMICA_EQUIPOS_PATIO = new MemoryStream(PANORAMICA_EQUIPOS_PATIO);
            byte[] BREAKER_ASIGNADO_PARA_SEGURIDAD = (byte[])dt.Rows[0]["BREAKER_ASIGNADO_PARA_SEGURIDAD"];
            MemoryStream mBREAKER_ASIGNADO_PARA_SEGURIDAD = new MemoryStream(BREAKER_ASIGNADO_PARA_SEGURIDAD);
            //byte[] CERRADURA_ELECTROMAGNETICA_EXTERNA = (byte[])ds.Tables[0].Rows[0]["CERRADURA_ELECTROMAGNETICA_EXTERNA"];
            //MemoryStream mCERRADURA_ELECTROMAGNETICA_EXTERNA = new MemoryStream(CERRADURA_ELECTROMAGNETICA_EXTERNA);
            byte[] CERRADURA_ELECTROMAGNETICA_EXTERNA2 = (byte[])dt.Rows[0]["CERRADURA_ELECTROMAGNETICA_EXTERNA2"];
            MemoryStream mCERRADURA_ELECTROMAGNETICA_EXTERNA2 = new MemoryStream(CERRADURA_ELECTROMAGNETICA_EXTERNA2);
            byte[] SENSOR_MAGNETICO_EXTERMO = (byte[])dt.Rows[0]["SENSOR_MAGNETICO_EXTERMO"];
            MemoryStream mSENSOR_MAGNETICO_EXTERMO = new MemoryStream(SENSOR_MAGNETICO_EXTERMO);
            byte[] SENSOR_MAGNETICO_EXTERNO2 = (byte[])dt.Rows[0]["SENSOR_MAGNETICO_EXTERNO2"];
            MemoryStream mSENSOR_MAGNETICO_EXTERNO2 = new MemoryStream(SENSOR_MAGNETICO_EXTERNO2);
            byte[] CERRADURA_ELECTRICA_EXTERNA = (byte[])dt.Rows[0]["CERRADURA_ELECTRICA_EXTERNA"];
            MemoryStream mCERRADURA_ELECTRICA_EXTERNA = new MemoryStream(CERRADURA_ELECTRICA_EXTERNA);
            byte[] SENSOR_MOVIMIENTO_90_EXTERNO_N1 = (byte[])dt.Rows[0]["SENSOR_MOVIMIENTO_90_EXTERNO_N1"];
            MemoryStream mSENSOR_MOVIMIENTO_90_EXTERNO_N1 = new MemoryStream(SENSOR_MOVIMIENTO_90_EXTERNO_N1);
            byte[] SENSOR_MOVIMIENTO_90_EXTERNO_N2 = (byte[])dt.Rows[0]["SENSOR_MOVIMIENTO_90_EXTERNO_N2"];
            MemoryStream mSENSOR_MOVIMIENTO_90_EXTERNO_N2 = new MemoryStream(SENSOR_MOVIMIENTO_90_EXTERNO_N2);
            byte[] SIRENA_ESTROBOSCOPICA = (byte[])dt.Rows[0]["SIRENA_ESTROBOSCOPICA"];
            MemoryStream mSIRENA_ESTROBOSCOPICA = new MemoryStream(SIRENA_ESTROBOSCOPICA);
            byte[] LECTOR_BIOMETRICO = (byte[])dt.Rows[0]["LECTOR_BIOMETRICO"];
            MemoryStream mLECTOR_BIOMETRICO = new MemoryStream(LECTOR_BIOMETRICO);
            byte[] LECTOR_TARJETA = (byte[])dt.Rows[0]["LECTOR_TARJETA"];
            MemoryStream mLECTOR_TARJETA = new MemoryStream(LECTOR_TARJETA);
            byte[] CAMARA_EXTERIOR_PTZ = (byte[])dt.Rows[0]["CAMARA_EXTERIOR_PTZ"];
            MemoryStream mCAMARA_EXTERIOR_PTZ = new MemoryStream(CAMARA_EXTERIOR_PTZ);
            byte[] EXTINTOR_EXTERIOR = (byte[])dt.Rows[0]["EXTINTOR_EXTERIOR"];
            MemoryStream mEXTINTOR_EXTERIOR = new MemoryStream(EXTINTOR_EXTERIOR);
            byte[] SENSOR_MAGNETICO_INTERNO = (byte[])dt.Rows[0]["SENSOR_MAGNETICO_INTERNO"];
            MemoryStream mSENSOR_MAGNETICO_INTERNO = new MemoryStream(SENSOR_MAGNETICO_INTERNO);
            byte[] SENSOR_MAGNETICO_INTERNO_2 = (byte[])dt.Rows[0]["SENSOR_MAGNETICO_INTERNO_2"];
            MemoryStream mSENSOR_MAGNETICO_INTERNO_2 = new MemoryStream(SENSOR_MAGNETICO_INTERNO_2);
            byte[] SENSOR_OCUPACIONAL = (byte[])dt.Rows[0]["SENSOR_OCUPACIONAL"];
            MemoryStream mSENSOR_OCUPACIONAL = new MemoryStream(SENSOR_OCUPACIONAL);
            byte[] SENSOR_DE_HUMO = (byte[])dt.Rows[0]["SENSOR_DE_HUMO"];
            MemoryStream mSENSOR_DE_HUMO = new MemoryStream(SENSOR_DE_HUMO);
            byte[] SENSOR_MOVIMIENTO_360 = (byte[])dt.Rows[0]["SENSOR_MOVIMIENTO_360"];
            MemoryStream mSENSOR_MOVIMIENTO_360 = new MemoryStream(SENSOR_MOVIMIENTO_360);
            byte[] SENSOR_DE_INUNDACION = (byte[])dt.Rows[0]["SENSOR_DE_INUNDACION"];
            MemoryStream mSENSOR_DE_INUNDACION = new MemoryStream(SENSOR_DE_INUNDACION);
            byte[] CAMARA_PTZ_INTERIOR = (byte[])dt.Rows[0]["CAMARA_PTZ_INTERIOR"];
            MemoryStream mCAMARA_PTZ_INTERIOR = new MemoryStream(CAMARA_PTZ_INTERIOR);
            byte[] EXTINTOR_INTERIOR = (byte[])dt.Rows[0]["EXTINTOR_INTERIOR"];
            MemoryStream mEXTINTOR_INTERIOR = new MemoryStream(EXTINTOR_INTERIOR);
            byte[] RELE_EQUIPO_INTERO = (byte[])dt.Rows[0]["RELE_EQUIPO_INTERO"];
            MemoryStream mRELE_EQUIPO_INTERO = new MemoryStream(RELE_EQUIPO_INTERO);
            byte[] CONTROLADOR_NVR_SWITCH = (byte[])dt.Rows[0]["CONTROLADOR_NVR_SWITCH"];
            MemoryStream mCONTROLADOR_NVR_SWITCH = new MemoryStream(CONTROLADOR_NVR_SWITCH);
            byte[] ATERRAMIENTO_CONTROLADOR = (byte[])dt.Rows[0]["ATERRAMIENTO_CONTROLADOR"];
            MemoryStream mATERRAMIENTO_CONTROLADOR = new MemoryStream(ATERRAMIENTO_CONTROLADOR);
            byte[] ATERRAMIENTO_NVR_POE = (byte[])dt.Rows[0]["ATERRAMIENTO_NVR_POE"];
            MemoryStream mATERRAMIENTO_NVR_POE = new MemoryStream(ATERRAMIENTO_NVR_POE);
            byte[] ATERRAMIENTO_NVR_POE_2 = (byte[])dt.Rows[0]["ATERRAMIENTO_NVR_POE_2"];
            MemoryStream mATERRAMIENTO_NVR_POE_2 = new MemoryStream(ATERRAMIENTO_NVR_POE_2);
            byte[] ATERRAMIENTO_A_BARRA = (byte[])dt.Rows[0]["ATERRAMIENTO_A_BARRA"];
            MemoryStream mATERRAMIENTO_A_BARRA = new MemoryStream(ATERRAMIENTO_A_BARRA);
            byte[] SERIAL_NUMBER_SENSOR_MOVIMIENTO_1 = (byte[])dt.Rows[0]["SERIAL_NUMBER_SENSOR_MOVIMIENTO_1"];
            MemoryStream mSERIAL_NUMBER_SENSOR_MOVIMIENTO_1 = new MemoryStream(SERIAL_NUMBER_SENSOR_MOVIMIENTO_1);
            byte[] SERIAL_NUMBER_SENSOR_MOVIMIENTO_2 = (byte[])dt.Rows[0]["SERIAL_NUMBER_SENSOR_MOVIMIENTO_2"];
            MemoryStream mSERIAL_NUMBER_SENSOR_MOVIMIENTO_2 = new MemoryStream(SERIAL_NUMBER_SENSOR_MOVIMIENTO_2);
            byte[] SERIAL_NUMBER_SWITCH_POE_NVR = (byte[])dt.Rows[0]["SERIAL_NUMBER_SWITCH_POE_NVR"];
            MemoryStream mSERIAL_NUMBER_SWITCH_POE_NVR = new MemoryStream(SERIAL_NUMBER_SWITCH_POE_NVR);
            byte[] SERIAL_NUMBER_SWITCH_POE_NVR_2 = (byte[])dt.Rows[0]["SERIAL_NUMBER_SWITCH_POE_NVR_2"];
            MemoryStream mSERIAL_NUMBER_SWITCH_POE_NVR_2 = new MemoryStream(SERIAL_NUMBER_SWITCH_POE_NVR_2);
            byte[] SERIAL_NUMBER_CONTROLADOR = (byte[])dt.Rows[0]["SERIAL_NUMBER_CONTROLADOR"];
            MemoryStream mSERIAL_NUMBER_CONTROLADOR = new MemoryStream(SERIAL_NUMBER_CONTROLADOR);
            byte[] ETIQUETADOS_EQUIPOS_CONTROLADOR = (byte[])dt.Rows[0]["ETIQUETADOS_EQUIPOS_CONTROLADOR"];
            MemoryStream mETIQUETADOS_EQUIPOS_CONTROLADOR = new MemoryStream(ETIQUETADOS_EQUIPOS_CONTROLADOR);
            byte[] ETIQUETADOS_EQUIPOS_NVR = (byte[])dt.Rows[0]["ETIQUETADOS_EQUIPOS_NVR"];
            MemoryStream mETIQUETADOS_EQUIPOS_NVR = new MemoryStream(ETIQUETADOS_EQUIPOS_NVR);
            byte[] CHECKLIST = (byte[])dt.Rows[0]["CHECKLIST"];
            MemoryStream mCHECKLIST = new MemoryStream(CHECKLIST);
            byte[] CAMARA_EXTERIOR_MODO_NORMAL_POS1 = (byte[])dt.Rows[0]["CAMARA_EXTERIOR_MODO_NORMAL_POS1"];
            MemoryStream mCAMARA_EXTERIOR_MODO_NORMAL_POS1 = new MemoryStream(CAMARA_EXTERIOR_MODO_NORMAL_POS1);
            byte[] CAMARA_EXTERIOR_MODO_NORMAL_POS2 = (byte[])dt.Rows[0]["CAMARA_EXTERIOR_MODO_NORMAL_POS2"];
            MemoryStream mCAMARA_EXTERIOR_MODO_NORMAL_POS2 = new MemoryStream(CAMARA_EXTERIOR_MODO_NORMAL_POS2);
            byte[] CAMARA_INTERIOR_MODO_NORMAL = (byte[])dt.Rows[0]["CAMARA_INTERIOR_MODO_NORMAL"];
            MemoryStream mCAMARA_INTERIOR_MODO_NORMAL = new MemoryStream(CAMARA_INTERIOR_MODO_NORMAL);
            byte[] CAMARA_INTERIOR_MODO_INFRARROJO = (byte[])dt.Rows[0]["CAMARA_INTERIOR_MODO_INFRARROJO"];
            MemoryStream mCAMARA_INTERIOR_MODO_INFRARROJO = new MemoryStream(CAMARA_INTERIOR_MODO_INFRARROJO);
            byte[] TPA_PUERTA_PRINCIPAL_ABIERTA = (byte[])dt.Rows[0]["TPA_PUERTA_PRINCIPAL_ABIERTA"];
            MemoryStream mTPA_PUERTA_PRINCIPAL_ABIERTA = new MemoryStream(TPA_PUERTA_PRINCIPAL_ABIERTA);
            byte[] TPA_PUERTA_SALAS_EQUIPOS_ABIERTA = (byte[])dt.Rows[0]["TPA_PUERTA_SALAS_EQUIPOS_ABIERTA"];
            MemoryStream mTPA_PUERTA_SALAS_EQUIPOS_ABIERTA = new MemoryStream(TPA_PUERTA_SALAS_EQUIPOS_ABIERTA);
            byte[] TPA_CAMARA_INTERNA = (byte[])dt.Rows[0]["TPA_CAMARA_INTERNA"];
            MemoryStream mTPA_CAMARA_INTERNA = new MemoryStream(TPA_CAMARA_INTERNA);
            byte[] TPA_CAMARA_EXTERNA = (byte[])dt.Rows[0]["TPA_CAMARA_EXTERNA"];
            MemoryStream mTPA_CAMARA_EXTERNA = new MemoryStream(TPA_CAMARA_EXTERNA);
            byte[] TPA_SENSOR_DE_ANIEGO = (byte[])dt.Rows[0]["TPA_SENSOR_DE_ANIEGO"];
            MemoryStream mTPA_SENSOR_DE_ANIEGO = new MemoryStream(TPA_SENSOR_DE_ANIEGO);
            byte[] TPA_SENSOR_DE_HUMO = (byte[])dt.Rows[0]["TPA_SENSOR_DE_HUMO"];
            MemoryStream mTPA_SENSOR_DE_HUMO = new MemoryStream(TPA_SENSOR_DE_HUMO);
            byte[] TPA_TAMPER_SENSOR_90_1 = (byte[])dt.Rows[0]["TPA_TAMPER_SENSOR_90_1"];
            MemoryStream mTPA_TAMPER_SENSOR_90_1 = new MemoryStream(TPA_TAMPER_SENSOR_90_1);
            byte[] TPA_MOVIMIENTO_SENSOR_90_1 = (byte[])dt.Rows[0]["TPA_MOVIMIENTO_SENSOR_90_1"];
            MemoryStream mTPA_MOVIMIENTO_SENSOR_90_1 = new MemoryStream(TPA_MOVIMIENTO_SENSOR_90_1);
            byte[] TPA_MASKING_SENSOR_90_1 = (byte[])dt.Rows[0]["TPA_MASKING_SENSOR_90_1"];
            MemoryStream mTPA_MASKING_SENSOR_90_1 = new MemoryStream(TPA_MASKING_SENSOR_90_1);
            byte[] TPA_TAMPER_SENSOR_90_2 = (byte[])dt.Rows[0]["TPA_TAMPER_SENSOR_90_2"];
            MemoryStream mTPA_TAMPER_SENSOR_90_2 = new MemoryStream(TPA_TAMPER_SENSOR_90_2);
            byte[] TPA_MOVIMIENTO_SENSOR_90_2 = (byte[])dt.Rows[0]["TPA_MOVIMIENTO_SENSOR_90_2"];
            MemoryStream mTPA_MOVIMIENTO_SENSOR_90_2 = new MemoryStream(TPA_MOVIMIENTO_SENSOR_90_2);
            byte[] TPA_MASKING_SENSOR_90_2 = (byte[])dt.Rows[0]["TPA_MASKING_SENSOR_90_2"];
            MemoryStream mTPA_MASKING_SENSOR_90_2 = new MemoryStream(TPA_MASKING_SENSOR_90_2);
            byte[] TPA_ALARMA_TAMPER_SENSOR_360 = (byte[])dt.Rows[0]["TPA_ALARMA_TAMPER_SENSOR_360"];
            MemoryStream mTPA_ALARMA_TAMPER_SENSOR_360 = new MemoryStream(TPA_ALARMA_TAMPER_SENSOR_360);
            byte[] TPA_ALARMA_MOVIMIENTO_SENSOR_360 = (byte[])dt.Rows[0]["TPA_ALARMA_MOVIMIENTO_SENSOR_360"];
            MemoryStream mTPA_ALARMA_MOVIMIENTO_SENSOR_360 = new MemoryStream(TPA_ALARMA_MOVIMIENTO_SENSOR_360);
            byte[] PING_CAMARA_1_INDOOR = (byte[])dt.Rows[0]["PING_CAMARA_1_INDOOR"];
            MemoryStream mPING_CAMARA_1_INDOOR = new MemoryStream(PING_CAMARA_1_INDOOR);
            byte[] PING_CAMARA_2_OUTDOOR = (byte[])dt.Rows[0]["PING_CAMARA_2_OUTDOOR"];
            MemoryStream mPING_CAMARA_2_OUTDOOR = new MemoryStream(PING_CAMARA_2_OUTDOOR);
            byte[] PING_CONTROLADOR = (byte[])dt.Rows[0]["PING_CONTROLADOR"];
            MemoryStream mPING_CONTROLADOR = new MemoryStream(PING_CONTROLADOR);
            byte[] PING_GATEWAY = (byte[])dt.Rows[0]["PING_GATEWAY"];
            MemoryStream mPING_GATEWAY = new MemoryStream(PING_GATEWAY);
            byte[] PING_NVR = (byte[])dt.Rows[0]["PING_NVR"];
            MemoryStream mPING_NVR = new MemoryStream(PING_NVR);
            byte[] PING_BIOMETRICO = (byte[])dt.Rows[0]["PING_BIOMETRICO"];
            MemoryStream mPING_BIOMETRICO = new MemoryStream(PING_BIOMETRICO);
            #endregion

            // String usuarioWindows = Environment.UserName;
            String excelGenerado = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";
            File.Copy(rutaPlantilla, excelGenerado, true);

            #region Agregando datos

            #region Caratula
            ExcelToolsBL.UpdateCell(excelGenerado, "Carátula", "NODO  " + NOMBRE_NODO, 15, "B");
            ExcelToolsBL.UpdateCell(excelGenerado, "Carátula", NOMBRE_NODO + " - PERU", 17, "B");
            ExcelToolsBL.UpdateCell(excelGenerado, "Carátula", TIPO_NODO, 22, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "Carátula", CODIGO_NODO, 24, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "Carátula", FECHA, 26, "D");
            #endregion

            #region Acta de Instalacion Aceptacion

            ExcelToolsBL.UpdateCell(excelGenerado, "Acta de Instal- aceptación", "NODO  " + NOMBRE_NODO, 15, "E");

            #endregion


            #region Reporte Fotografico

            //aumentar ancho 70 y alto 60

            ExcelToolsBL.UpdateCell(excelGenerado, "Reporte fotográfico", " SISTEMAS DE SEGURIDAD NODO " + TIPO_NODO + "_" + NOMBRE_NODO, 5, "B");

            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mFACHADA_DEL_NODO, "", 10,3,406,263);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mSALA_EQUIPOS_PANORAMICA_RACK, "", 10,14, 442, 262);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mPANORAMICA_INTERIOR_01, "", 26, 3,408, 261);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mPANORAMICA_INTERIOR_02, "", 26, 14, 444, 262);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mPANORAMICA_EQUIPOS_PATIO, "", 42, 3, 408, 261);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mBREAKER_ASIGNADO_PARA_SEGURIDAD, "", 42, 14, 443, 262);
            //ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mCERRADURA_ELECTROMAGNETICA_EXTERNA, "", 59, 3, 407, 240);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mCERRADURA_ELECTROMAGNETICA_EXTERNA2, "", 59,3, 407, 248);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mSENSOR_MAGNETICO_EXTERMO, "", 59, 14, 226, 241);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mSENSOR_MAGNETICO_EXTERNO2, "", 59, 19, 216, 240);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mCERRADURA_ELECTRICA_EXTERNA, "", 78,3, 408, 305);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mSENSOR_MOVIMIENTO_90_EXTERNO_N1, "", 78, 14, 442, 304);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mSENSOR_MOVIMIENTO_90_EXTERNO_N2, "", 96, 3, 409, 267);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mSIRENA_ESTROBOSCOPICA, "", 96, 14,442,266);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mLECTOR_BIOMETRICO, "", 116,3,406,295);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mLECTOR_TARJETA, "", 116, 14, 442, 295);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mCAMARA_EXTERIOR_PTZ, "", 135, 3, 408, 310);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mEXTINTOR_EXTERIOR, "", 135, 14, 442, 309);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mSENSOR_MAGNETICO_INTERNO, "", 156, 3, 181, 278);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mSENSOR_MAGNETICO_INTERNO_2, "", 156, 7, 226, 278);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mSENSOR_OCUPACIONAL, "", 156, 14, 442, 278);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mSENSOR_DE_HUMO, "", 172, 3, 406, 312);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mSENSOR_MOVIMIENTO_360, "", 172, 14,442,313);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mSENSOR_DE_INUNDACION, "", 188,3,406, 311);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mCAMARA_PTZ_INTERIOR, "", 188, 14, 442,311);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mEXTINTOR_INTERIOR, "", 205,3, 407, 464);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mRELE_EQUIPO_INTERO, "", 205, 14, 442,463);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mCONTROLADOR_NVR_SWITCH, "", 222,3,894,338);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mATERRAMIENTO_CONTROLADOR, "", 247, 3,406, 304);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mATERRAMIENTO_NVR_POE, "", 247, 14, 227,305);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mATERRAMIENTO_NVR_POE_2, "", 247, 19, 217, 305);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mATERRAMIENTO_A_BARRA, "", 263, 3,408,272);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mSERIAL_NUMBER_SENSOR_MOVIMIENTO_1, "", 282,3,407,278);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mSERIAL_NUMBER_SENSOR_MOVIMIENTO_2, "", 282,14, 446,278);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mSERIAL_NUMBER_SWITCH_POE_NVR, "", 298,3,183,332);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mSERIAL_NUMBER_SWITCH_POE_NVR_2, "", 298,7,227,330);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mSERIAL_NUMBER_CONTROLADOR, "", 298,14,442,332);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mETIQUETADOS_EQUIPOS_CONTROLADOR, "", 317,5,678,312);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mETIQUETADOS_EQUIPOS_NVR, "", 335,5,679,311);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mCHECKLIST, "", 353, 7, 435, 542);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mCAMARA_EXTERIOR_MODO_NORMAL_POS1, "",391,5,722, 370);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mCAMARA_EXTERIOR_MODO_NORMAL_POS2, "",408,5,723, 312);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mCAMARA_INTERIOR_MODO_NORMAL, "", 425,5,722,312);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mCAMARA_INTERIOR_MODO_INFRARROJO, "", 442,5,677, 341);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mTPA_PUERTA_PRINCIPAL_ABIERTA, "", 459,5,724,378);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mTPA_PUERTA_SALAS_EQUIPOS_ABIERTA, "", 476,5,723,378);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mTPA_CAMARA_INTERNA, "", 492,5, 723, 378);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mTPA_CAMARA_EXTERNA, "", 507,5,723,378);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mTPA_SENSOR_DE_ANIEGO, "", 522, 5,723,378);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mTPA_SENSOR_DE_HUMO, "", 539,5,723,378);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mTPA_TAMPER_SENSOR_90_1, "", 554,5,722,373);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mTPA_MOVIMIENTO_SENSOR_90_1, "", 569,5,723,372);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mTPA_MASKING_SENSOR_90_1, "", 584,5,722,371);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mTPA_TAMPER_SENSOR_90_2, "", 599,5,723,372);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mTPA_MOVIMIENTO_SENSOR_90_2, "",614,5,725,373);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mTPA_MASKING_SENSOR_90_2, "", 629,5,722,373);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mTPA_ALARMA_TAMPER_SENSOR_360, "", 644,5,724,372);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mTPA_ALARMA_MOVIMIENTO_SENSOR_360, "",659,5,722, 371);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mPING_CAMARA_1_INDOOR, "",676, 5,723,310);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mPING_CAMARA_2_OUTDOOR, "", 693,5,724,310);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mPING_CONTROLADOR, "", 710, 5,722,312);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mPING_GATEWAY, "", 727, 5,723,359);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mPING_NVR, "", 744,5,723,337);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mPING_BIOMETRICO, "", 761,5,723,338);
            #endregion

            #region Materiales

            ExcelToolsBL.UpdateCell(excelGenerado, "Materiales", "NODO  " + TIPO_NODO, 11, "C");
            ExcelToolsBL.UpdateCell(excelGenerado, "Materiales", CODIGO_NODO, 11, "F");


            foreach (DataRow dr in dt1.Rows)
            {
                String EQUIPO = dr["EQUIPO"].ToString();
                String MARCA = dr["MARCA"].ToString();
                String MODELO = dr["MODELO"].ToString();
                String Nro_SERIE = dr["Nro_SERIE"].ToString();

                int ind = Convert.ToInt32(dt1.Rows.IndexOf(dr));

                ExcelToolsBL.UpdateCell(excelGenerado, "Materiales", Convert.ToString(ind + 1), 16 + ind, "B");
                ExcelToolsBL.UpdateCell(excelGenerado, "Materiales", EQUIPO, 16 + ind, "C");
                ExcelToolsBL.UpdateCell(excelGenerado, "Materiales", "1", 16 + ind, "D");
                ExcelToolsBL.UpdateCell(excelGenerado, "Materiales", MARCA, 16 + ind, "E");
                ExcelToolsBL.UpdateCell(excelGenerado, "Materiales", MODELO, 16 + ind, "F");
                ExcelToolsBL.UpdateCell(excelGenerado, "Materiales", Nro_SERIE, 16 + ind, "G");
            }


            foreach (DataRow dr in dt2.Rows)
            {
                String DESCRIPCION = dr["DESCRIPCION"].ToString();
                String CODIGO = dr["CODIGO"].ToString();
                String UNIDAD = dr["UNIDAD"].ToString();
                String CANTIDAD = dr["CANTIDAD"].ToString();

                int ind = Convert.ToInt32(dt2.Rows.IndexOf(dr));

                ExcelToolsBL.UpdateCell(excelGenerado, "Materiales", Convert.ToString(ind + 1), 40 + ind, "B");
                ExcelToolsBL.UpdateCell(excelGenerado, "Materiales", DESCRIPCION, 40 + ind, "C");
                ExcelToolsBL.UpdateCell(excelGenerado, "Materiales", UNIDAD, 40 + ind, "E");
                ExcelToolsBL.UpdateCell(excelGenerado, "Materiales", CANTIDAD, 40 + ind, "F");

            }

            #endregion

            #region ATP
            ExcelToolsBL.UpdateCell(excelGenerado, "ATP", DEPARTAMENTO,9,"C");
            ExcelToolsBL.UpdateCell(excelGenerado, "ATP", FECHA,6,"C");
            ExcelToolsBL.UpdateCell(excelGenerado, "ATP", CODIGO_NODO,9,"J");
            ExcelToolsBL.UpdateCell(excelGenerado, "ATP", NOMBRE_NODO,10,"J");
            ExcelToolsBL.UpdateCell(excelGenerado, "ATP", EXTINGUIDOR_EXT_FECHA_EXPIRACION, 43, "G");
            ExcelToolsBL.UpdateCell(excelGenerado, "ATP", EXTINGUIDOR_INT_FECHA_EXPIRACION, 44, "G");
            ExcelToolsBL.UpdateCell(excelGenerado, "ATP", SERIAL_CONTROLADOR, 8,"C");
            //IP CONTROLADOR
            ExcelToolsBL.UpdateCell(excelGenerado, "ATP", IP_CONTROLADOR, 13, "C");
            #endregion

            #endregion

            #region Ruta del Zip
            String rutaNodo = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo;
            if (Directory.Exists(rutaNodo))
            {
                String rutaParcial = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\TRANSPORTE\\";
                if (!Directory.Exists(rutaParcial)) Directory.CreateDirectory(rutaParcial);
                String rutaAlterna = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\TRANSPORTE\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";
                File.Copy(excelGenerado, rutaAlterna, true);
            }
            #endregion
        }

        public void EstudioDeCampo(String IdNodo, String IdTarea, String valorCadena1, String rutaPlantilla)
        {
            DBBaseDatos baseDatosDA = new DBBaseDatos();

            baseDatosDA.Configurar();
            baseDatosDA.Conectar();
            DataTable dt = new DataTable();
            try
            {
                baseDatosDA.CrearComando("USP_R_ESTUDIO_DE_CAMPO", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                dt = baseDatosDA.EjecutarConsultaDataTable();


              
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
            #region Ingresando Strings

            String fechaSQL = dt.Rows[0]["FECHA"].ToString();
            String FECHA = "";
            if (fechaSQL != "")
            {
                DateTime dtFecha = DateTime.Parse(fechaSQL);
                FECHA = dtFecha.ToString("dd/MM/yyyy");
            }
            else { FECHA = ""; }

            // String HORA_INICIO = dt.Rows[0]["HORA_INICIO"].ToString();
            //String HORA_FIN = dt.Rows[0]["HORA_FIN"].ToString();
            String TIPO_NODO = dt.Rows[0]["TIPO_NODO"].ToString();
            String UBIGEO = dt.Rows[0]["UBIGEO"].ToString();
            String DEPARTAMENTO = dt.Rows[0]["DEPARTAMENTO"].ToString();
            String PROVINCIA = dt.Rows[0]["PROVINCIA"].ToString();
            String DISTRITO = dt.Rows[0]["DISTRITO"].ToString();
            String NOMBRE_NODO = dt.Rows[0]["NOMBRE_NODO"].ToString();
            String LONGITUD_LOCALIDAD_PLAZA_PRINCIPAL = dt.Rows[0]["LONGITUD_LOCALIDAD_PLAZA_PRINCIPAL"].ToString();
            String LATITUD_LOCALIDAD_PLAZA_PRINCIPAL = dt.Rows[0]["LATITUD_LOCALIDAD_PLAZA_PRINCIPAL"].ToString();
            String ALTURA_MSNM = dt.Rows[0]["ALTURA_MSNM"].ToString();
            String AREA_NATURAL_PROTEG = dt.Rows[0]["AREA_NATURAL_PROTEG"].ToString();
            String NOMBRE_AREA_NATURAL = dt.Rows[0]["NOMBRE_AREA_NATURAL"].ToString();
            String RESTOS_ARQUEOLOGICOS = dt.Rows[0]["RESTOS_ARQUEOLOGICOS"].ToString();
            String ESPECIF_TIPO_RESTOS_ARQ = dt.Rows[0]["ESPECIF_TIPO_RESTOS_ARQ"].ToString();
            String BANCO_NACION = dt.Rows[0]["BANCO_NACION"].ToString();
            String AGENTE_BANCO_NACION = dt.Rows[0]["AGENTE_BANCO_NACION"].ToString();
            String CANTIDAD = dt.Rows[0]["CANTIDAD"].ToString();
            String OTROS_BANCOS = dt.Rows[0]["OTROS_BANCOS"].ToString();
            String CANTIDAD_OTROS_BANCOS = dt.Rows[0]["CANTIDAD_OTROS_BANCOS"].ToString();
            String ENTIDAD_IMPORTANTE = dt.Rows[0]["ENTIDAD_IMPORTANTE"].ToString();
            String IIEE = dt.Rows[0]["IIEE"].ToString();
            String CANTIDAD_IIEE = dt.Rows[0]["CANTIDAD_IIEE"].ToString();
            String POBLACION = dt.Rows[0]["POBLACION"].ToString();
            String N_DE_MUJERES = dt.Rows[0]["N_DE_MUJERES"].ToString();
            String N_DE_JOVENES_15_24 = dt.Rows[0]["N_DE_JOVENES_15_24"].ToString();
            String N_DE_PERSONAS_DISCAPACIDAD = dt.Rows[0]["N_DE_PERSONAS_DISCAPACIDAD"].ToString();
            String N_VIVIENDAS = dt.Rows[0]["N_VIVIENDAS"].ToString();
            // String ENTIDAD_IMPORTANTE_2 = dt.Rows[0]["ENTIDAD_IMPORTANTE_2"].ToString(); //NO EXISTE EN EL EXCEL LUGAR DONDE COLOCARLO, PENDIENTE
            #endregion

            // String usuarioWindows = Environment.UserName;
            String excelGenerado = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";

            File.Copy(rutaPlantilla, excelGenerado, true);

            #region Agregando datos al excel
            ExcelToolsBL.UpdateCell(excelGenerado, "Sheet3", FECHA, 2, "A");
            ExcelToolsBL.UpdateCell(excelGenerado, "Sheet3", "08:00", 2, "B");
            ExcelToolsBL.UpdateCell(excelGenerado, "Sheet3", "18:00", 2, "C");
            ExcelToolsBL.UpdateCell(excelGenerado, "Sheet3", TIPO_NODO, 2, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "Sheet3", UBIGEO, 2, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "Sheet3", DEPARTAMENTO, 2, "F");
            ExcelToolsBL.UpdateCell(excelGenerado, "Sheet3", PROVINCIA, 2, "G");
            ExcelToolsBL.UpdateCell(excelGenerado, "Sheet3", DISTRITO, 2, "H");
            ExcelToolsBL.UpdateCell(excelGenerado, "Sheet3", NOMBRE_NODO, 2, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "Sheet3", LONGITUD_LOCALIDAD_PLAZA_PRINCIPAL, 2, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "Sheet3", LATITUD_LOCALIDAD_PLAZA_PRINCIPAL, 2, "K");
            ExcelToolsBL.UpdateCell(excelGenerado, "Sheet3", ALTURA_MSNM, 2, "L");
            ExcelToolsBL.UpdateCell(excelGenerado, "Sheet3", AREA_NATURAL_PROTEG, 2, "M");
            ExcelToolsBL.UpdateCell(excelGenerado, "Sheet3", NOMBRE_AREA_NATURAL, 2, "N");
            ExcelToolsBL.UpdateCell(excelGenerado, "Sheet3", RESTOS_ARQUEOLOGICOS, 2, "O");
            ExcelToolsBL.UpdateCell(excelGenerado, "Sheet3", ESPECIF_TIPO_RESTOS_ARQ, 2, "P");
            ExcelToolsBL.UpdateCell(excelGenerado, "Sheet3", BANCO_NACION, 2, "Q");
            ExcelToolsBL.UpdateCell(excelGenerado, "Sheet3", AGENTE_BANCO_NACION, 2, "R");
            ExcelToolsBL.UpdateCell(excelGenerado, "Sheet3", CANTIDAD, 2, "S");
            ExcelToolsBL.UpdateCell(excelGenerado, "Sheet3", OTROS_BANCOS, 2, "T");
            ExcelToolsBL.UpdateCell(excelGenerado, "Sheet3", CANTIDAD_OTROS_BANCOS, 2, "U");
            ExcelToolsBL.UpdateCell(excelGenerado, "Sheet3", ENTIDAD_IMPORTANTE, 2, "V");
            ExcelToolsBL.UpdateCell(excelGenerado, "Sheet3", IIEE, 2, "W");
            ExcelToolsBL.UpdateCell(excelGenerado, "Sheet3", CANTIDAD_IIEE, 2, "X");
            ExcelToolsBL.UpdateCell(excelGenerado, "Sheet3", POBLACION, 2, "Y");
            ExcelToolsBL.UpdateCell(excelGenerado, "Sheet3", N_DE_MUJERES, 2, "Z");
            ExcelToolsBL.UpdateCell(excelGenerado, "Sheet3", N_DE_JOVENES_15_24, 2, "AA");
            ExcelToolsBL.UpdateCell(excelGenerado, "Sheet3", N_DE_PERSONAS_DISCAPACIDAD, 2, "AB");
            ExcelToolsBL.UpdateCell(excelGenerado, "Sheet3", N_VIVIENDAS, 2, "AC");
            #endregion

            #region Ruta del Zip
            String rutaNodo = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo;
            if (Directory.Exists(rutaNodo))
            {
                String rutaParcial = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\ESTUDIO DE CAMPO\\";
                if (!Directory.Exists(rutaParcial)) Directory.CreateDirectory(rutaParcial);
                String rutaAlterna = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\ESTUDIO DE CAMPO\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";
                File.Copy(excelGenerado, rutaAlterna, true);
            }
            #endregion
        }

        public void ProtocoloInstalacion(String IdNodo, String IdTarea, String valorCadena1, String rutaPlantilla)
        {
            DBBaseDatos baseDatosDA = new DBBaseDatos();

            baseDatosDA.Configurar();
            baseDatosDA.Conectar();
            DataTable dt = new DataTable();

            try
            {
                baseDatosDA.CrearComando("USP_R_PROTOCOLO_INSTALACION",CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                dt = baseDatosDA.EjecutarConsultaDataTable();

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
            #region Valores

            String DEPARTAMENTO = dt.Rows[0]["DEPARTAMENTO"].ToString();
            String NOMBRE_NODO = dt.Rows[0]["NOMBRE_NODO"].ToString();
            String TIPO_NODO = dt.Rows[0]["TIPO_NODO"].ToString();
            String CODIGO_NODO = dt.Rows[0]["CODIGO_NODO"].ToString();

            String fechaSQL = dt.Rows[0]["FECHA"].ToString();
            String FECHA = "";
            if (fechaSQL != "")
            {
                DateTime dtFecha = DateTime.Parse(fechaSQL);
                FECHA = dtFecha.ToString("dd/MM/yyyy");
            }
            else { FECHA = ""; }

            String NUM_SERIE_SWITCH = dt.Rows[0]["NUM_SERIE_SWITCH"].ToString();

            byte[] OMNISWITCH = (byte[])dt.Rows[0]["FOTO_1_OMNISWITCH"];
            MemoryStream mOMNISWITCH = new MemoryStream(OMNISWITCH);
            byte[] PAN_RACK = (byte[])dt.Rows[0]["FOTO_2_PAN_RACK"];
            MemoryStream mPAN_RACK = new MemoryStream(PAN_RACK);
            byte[] CON_BREAKERS_ASIGNADOS = (byte[])dt.Rows[0]["FOTO_3_CON_BREAKERS_ASIGNADOS"];
            MemoryStream mCON_BREAKERS_ASIGNADOS = new MemoryStream(CON_BREAKERS_ASIGNADOS);
            byte[] CON_ALIMEN_SWITCH = (byte[])dt.Rows[0]["FOTO_4_CON_ALIMEN_SWITCH"];
            MemoryStream mCON_ALIMEN_SWITCH = new MemoryStream(CON_ALIMEN_SWITCH);
            byte[] ATERRAMIENTO_SWITCH = (byte[])dt.Rows[0]["FOTO_5_ATERRAMIENTO_SWITCH"];
            MemoryStream mATERRAMIENTO_SWITCH = new MemoryStream(ATERRAMIENTO_SWITCH);
            byte[] ATERRAMIENTO_BARRA = (byte[])dt.Rows[0]["FOTO_6_ATERRAMIENTO_BARRA"];
            MemoryStream mATERRAMIENTO_BARRA = new MemoryStream(ATERRAMIENTO_BARRA);

            #endregion

            // String usuarioWindows = Environment.UserName;
            String excelGenerado = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";

            File.Copy(rutaPlantilla, excelGenerado, true);

            #region Agregando Valores por Hoja Excel

            #region Caratula 

            ExcelToolsBL.UpdateCell(excelGenerado, "Carátula", "NODO: " + " " + CODIGO_NODO + " " + " " + NOMBRE_NODO, 15, "B");
            ExcelToolsBL.UpdateCell(excelGenerado, "Carátula", DEPARTAMENTO + "-PERU", 17, "B");
            ExcelToolsBL.UpdateCell(excelGenerado, "Carátula", TIPO_NODO, 22, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "Carátula", CODIGO_NODO, 24, "D");
            //areglar el formato de fecha
            ExcelToolsBL.UpdateCell(excelGenerado, "Carátula", FECHA, 26, "D");
            #endregion

            ExcelToolsBL.UpdateCell(excelGenerado, "Acta de Instal- aceptación", "NODO: " + CODIGO_NODO, 15, "E");

            #region Reporte Fotografico
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mOMNISWITCH, "", 11, 5,675,312);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mPAN_RACK, "",29,5,675,312);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mCON_BREAKERS_ASIGNADOS, "", 46, 5,675,312);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mCON_ALIMEN_SWITCH, "", 62, 5,675,312);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mATERRAMIENTO_SWITCH, "", 80, 5,675,312);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mATERRAMIENTO_BARRA, "", 97, 5,675,312);
            #endregion

            #region Materiales
            ExcelToolsBL.UpdateCell(excelGenerado, "Materiales", "NODO: " + NOMBRE_NODO, 11, "C");
            ExcelToolsBL.UpdateCell(excelGenerado, "Materiales", CODIGO_NODO, 11, "F");
            ExcelToolsBL.UpdateCell(excelGenerado, "Materiales", NUM_SERIE_SWITCH, 17, "G");
            #endregion

            #endregion

            #region Ruta del Zip
            String rutaNodo = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo;
            if (Directory.Exists(rutaNodo))
            {
                String rutaParcial = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\ACCESO\\PROTOCOLO_INSTALACION\\";
                if (!Directory.Exists(rutaParcial)) Directory.CreateDirectory(rutaParcial);
                String rutaAlterna = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\ACCESO\\PROTOCOLO_INSTALACION\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";
                File.Copy(excelGenerado, rutaAlterna, true);
            }
            #endregion
        }

        public void ActaInstalacionAceptacionProtocoloSectorial(String IdNodo, String IdTarea, String valorCadena1, String rutaPlantilla)
        {
            DBBaseDatos baseDatosDA = new DBBaseDatos();

            baseDatosDA.Configurar();
            baseDatosDA.Conectar();
            DataTable dt = new DataTable();
            DataTable dt1 = new DataTable();
            DataTable dt2 = new DataTable();
            DataTable dt3 = new DataTable();
            try
            {
                baseDatosDA.CrearComando("USP_R_ACTA_INSTALACION_ACEPTACION_PROTOCOLO_SECTORIAL", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                dt = baseDatosDA.EjecutarConsultaDataTable();

                baseDatosDA.CrearComando("USP_R_MEDICION_SECTORIAL", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                dt1 = baseDatosDA.EjecutarConsultaDataTable();

                baseDatosDA.CrearComando("USP_R_IIBB_POR_TAREA", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                dt2 = baseDatosDA.EjecutarConsultaDataTable();

                baseDatosDA.CrearComando("USP_R_MATERIALES_PROTOCOLO_SECTORIAL", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                dt3 = baseDatosDA.EjecutarConsultaDataTable();

              
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
            #region Valores


            #region Valores String

            String NOMBRE_NODO = dt.Rows[0]["NOMBRE_NODO"].ToString();
            String CODIGO_NODO = dt.Rows[0]["CODIGO_NODO"].ToString();
            String TIPO_NODO = dt.Rows[0]["TIPO_NODO"].ToString();
            String FRECUENCIA = dt.Rows[0]["FRECUENCIA"].ToString();

            String fechaSQL = dt.Rows[0]["FECHA"].ToString();
            String FECHA = "";
            if (fechaSQL != "")
            {
                DateTime dtFecha = DateTime.Parse(fechaSQL);
                FECHA = dtFecha.ToString("dd/MM/yyyy");
            }
            else { FECHA = ""; }

            String DEPARTAMENTO = dt.Rows[0]["DEPARTAMENTO"].ToString();
            String PROVINCIA = dt.Rows[0]["PROVINCIA"].ToString();
            String DISTRITO = dt.Rows[0]["DISTRITO"].ToString();
            String DIRECCION = dt.Rows[0]["DIRECCION"].ToString();
            String LATITUD_S = dt.Rows[0]["LATITUD_S"].ToString();
            String LONGITUD_W = dt.Rows[0]["LONGITUD_W"].ToString();
            String ALTITUD_MSNM = dt.Rows[0]["ALTITUD_MSNM"].ToString();
            String AZIMUT = dt.Rows[0]["AZIMUT"].ToString();
            String DOWN_TILT = dt.Rows[0]["DOWN_TILT"].ToString();
            String ALTURA_INST = dt.Rows[0]["ALTURA_INST"].ToString();
            String IP_ADDRESS = dt.Rows[0]["IP_ADDRESS"].ToString();
            String GATEWAY_IP = dt.Rows[0]["GATEWAY_IP"].ToString();
            String CAPACIDAD_ENLACE = dt.Rows[0]["CAPACIDAD_ENLACE"].ToString();
            String EFICIENCIA_ENLACE = dt.Rows[0]["EFICIENCIA_ENLACE"].ToString();
            String SITE_NAME_SSID = dt.Rows[0]["SITE_NAME_SSID"].ToString();

            //String DISTANCIA_B = dt.Rows[0]["DISTANCIA_B"].ToString();
            //String DISTANCIA_C = dt.Rows[0]["DISTANCIA_C"].ToString();
            //String DISTANCIA_D = dt.Rows[0]["DISTANCIA_D"].ToString();
            //String DISTANCIA_E = dt.Rows[0]["DISTANCIA_E"].ToString();
            String ALTURA_TORRE = dt.Rows[0]["ALTURA_TORRE"].ToString();

            String SERIE_ODU = dt.Rows[0]["#SERIE_ODU"].ToString();
            String SERIE_ANTENA = dt.Rows[0]["#SERIE_ANTENA"].ToString();

            String SERIAL_TERMINAL_ACCESS_POINT = dt.Rows[0]["SERIAL_TERMINAL_ACCESS_POINT"].ToString();
            String SERIAL_ANTENA_CAMBIUM_5GHZ = dt.Rows[0]["SERIAL_ANTENA_CAMBIUM_5GHZ"].ToString();
            String SERIAL_POE_INYECTOR = dt.Rows[0]["SERIAL_POE_INYECTOR"].ToString();

            String UBIGEO = dt.Rows[0]["UBIGEO"].ToString();

            String MODELO_PUERTO_NODO = dt.Rows[0]["MODELO_PUERTO_NODO"].ToString();
            String PUERTO_NODO = dt.Rows[0]["PUERTO_NODO"].ToString();

            /*   int a = Convert.ToInt32(ALTURA_INST);
         int b = Convert.ToInt32(DISTANCIA_B);
         int c = Convert.ToInt32(DISTANCIA_C);
         int d = Convert.ToInt32(DISTANCIA_D);
         int e = Convert.ToInt32(DISTANCIA_E);

         int L = a + b - c + d + e;

         int LT = L + 3;  */

            #endregion

            #region Valores Binarios

            byte[] CAP1_CONF_RADIO = (byte[])dt.Rows[0]["CAP1_CONF_RADIO"];
            MemoryStream mCAP1_CONF_RADIO = new MemoryStream(CAP1_CONF_RADIO);
            byte[] CAP2_CONF_QoS = (byte[])dt.Rows[0]["CAP2_CONF_QoS"];
            MemoryStream mCAP2_CONF_QoS = new MemoryStream(CAP2_CONF_QoS);
            byte[] CAP3_1_CONF_SYSTEM = (byte[])dt.Rows[0]["CAP3_1_CONF_SYSTEM"];
            MemoryStream mCAP3_1_CONF_SYSTEM = new MemoryStream(CAP3_1_CONF_SYSTEM);
            byte[] CAP3_2_CONF_SYSTEM = (byte[])dt.Rows[0]["CAP3_2_CONF_SYSTEM"];
            MemoryStream mCAP3_2_CONF_SYSTEM = new MemoryStream(CAP3_2_CONF_SYSTEM);
            byte[] CAP4_MONITOR_SYSTEM = (byte[])dt.Rows[0]["CAP4_MONITOR_SYSTEM"];
            MemoryStream mCAP4_MONITOR_SYSTEM = new MemoryStream(CAP4_MONITOR_SYSTEM);
            byte[] CAP5_1_MON_WIRELESS = (byte[])dt.Rows[0]["CAP5_1_MON_WIRELESS"];
            MemoryStream mCAP5_1_MON_WIRELESS = new MemoryStream(CAP5_1_MON_WIRELESS);
            byte[] CAP5_2_MON_WIRELESS = (byte[])dt.Rows[0]["CAP5_2_MON_WIRELESS"];
            MemoryStream mCAP5_2_MON_WIRELESS = new MemoryStream(CAP5_2_MON_WIRELESS);
            byte[] CAP6_TOOLS_WIRELESS = (byte[])dt.Rows[0]["CAP6_TOOLS_WIRELESS"];
            MemoryStream mCAP6_TOOLS_WIRELESS = new MemoryStream(CAP6_TOOLS_WIRELESS);
            byte[] CAP7_PANTALLA_HOME = (byte[])dt.Rows[0]["CAP7_PANTALLA_HOME"];
            MemoryStream mCAP7_PANTALLA_HOME = new MemoryStream(CAP7_PANTALLA_HOME);

            byte[] FOTO1_PAN_ESTACION_A = (byte[])dt.Rows[0]["FOTO1_PAN_ESTACION_A"];
            MemoryStream mFOTO1_PAN_ESTACION_A = new MemoryStream(FOTO1_PAN_ESTACION_A);
            byte[] FOTO2_P0S_ANTENA_INST = (byte[])dt.Rows[0]["FOTO2_P0S_ANTENA_INST"];
            MemoryStream mFOTO2_P0S_ANTENA_INST = new MemoryStream(FOTO2_P0S_ANTENA_INST);
            byte[] FOTO3_ANTENA_ODU_TORRE = (byte[])dt.Rows[0]["FOTO3_ANTENA_ODU_TORRE"];
            MemoryStream mFOTO3_ANTENA_ODU_TORRE = new MemoryStream(FOTO3_ANTENA_ODU_TORRE);
            byte[] FOTO4_UGPS = (byte[])dt.Rows[0]["FOTO4_UGPS"];
            MemoryStream mFOTO4_UGPS = new MemoryStream(FOTO4_UGPS);
            byte[] FOTO5_ENGRASADO_PERNO = (byte[])dt.Rows[0]["FOTO5_ENGRASADO_PERNO"];
            MemoryStream mFOTO5_ENGRASADO_PERNO = new MemoryStream(FOTO5_ENGRASADO_PERNO);
            byte[] FOTO6_SILICONEADO_ETIQUETADO = (byte[])dt.Rows[0]["FOTO6_SILICONEADO_ETIQUETADO"];
            MemoryStream mFOTO6_SILICONEADO_ETIQUETADO = new MemoryStream(FOTO6_SILICONEADO_ETIQUETADO);
            byte[] FOTO8_RECORRIDO_CABLE_SFTP = (byte[])dt.Rows[0]["FOTO8_RECORRIDO_CABLE_SFTP"];
            MemoryStream mFOTO8_RECORRIDO_CABLE_SFTP = new MemoryStream(FOTO8_RECORRIDO_CABLE_SFTP);
            byte[] FOTO9_ATERRAMIENTO_CABLE_SFTP_OUT = (byte[])dt.Rows[0]["FOTO9_ATERRAMIENTO_CABLE_SFTP_OUT"];
            MemoryStream mFOTO9_ATERRAMIENTO_CABLE_SFTP_OUT = new MemoryStream(FOTO9_ATERRAMIENTO_CABLE_SFTP_OUT);
            byte[] FOTO10_ATERRAMIENTO_CABLE_SFTP_IN = (byte[])dt.Rows[0]["FOTO10_ATERRAMIENTO_CABLE_SFTP_IN"];
            MemoryStream mFOTO10_ATERRAMIENTO_CABLE_SFTP_IN = new MemoryStream(FOTO10_ATERRAMIENTO_CABLE_SFTP_IN);
            byte[] FOTO11_ETIQUETADO_POE = (byte[])dt.Rows[0]["FOTO11_ETIQUETADO_POE"];
            MemoryStream mFOTO11_ETIQUETADO_POE = new MemoryStream(FOTO11_ETIQUETADO_POE);
            byte[] FOTO12_PAN_RACK = (byte[])dt.Rows[0]["FOTO12_PAN_RACK"];
            MemoryStream mFOTO12_PAN_RACK = new MemoryStream(FOTO12_PAN_RACK);
            byte[] FOTO13_ATERRAMIENTO_POE = (byte[])dt.Rows[0]["FOTO13_ATERRAMIENTO_POE"];
            MemoryStream mFOTO13_ATERRAMIENTO_POE = new MemoryStream(FOTO13_ATERRAMIENTO_POE);
            byte[] FOTO14_1_EMERGENCIA_POE_ETIQUETA = (byte[])dt.Rows[0]["FOTO14_1_EMERGENCIA_POE_ETIQUETA"];
            MemoryStream mFOTO14_1_EMERGENCIA_POE_ETIQUETA = new MemoryStream(FOTO14_1_EMERGENCIA_POE_ETIQUETA);
            byte[] FOTO14_2_EMERGENCIA_POE_ETIQUETA = (byte[])dt.Rows[0]["FOTO14_2_EMERGENCIA_POE_ETIQUETA"];
            MemoryStream mFOTO14_2_EMERGENCIA_POE_ETIQUETA = new MemoryStream(FOTO14_2_EMERGENCIA_POE_ETIQUETA);
            byte[] FOTO15_PATCH_CORE_SALIDA_POE = (byte[])dt.Rows[0]["FOTO15_PATCH_CORE_SALIDA_POE"];
            MemoryStream mFOTO15_PATCH_CORE_SALIDA_POE = new MemoryStream(FOTO15_PATCH_CORE_SALIDA_POE);
            byte[] FOTO16_PATCH_CORE_SALIDA_SWITCH = (byte[])dt.Rows[0]["FOTO16_PATCH_CORE_SALIDA_SWITCH"];
            MemoryStream mFOTO16_PATCH_CORE_SALIDA_SWITCH = new MemoryStream(FOTO16_PATCH_CORE_SALIDA_SWITCH);
            byte[] FOTO17_SERIE_POE = (byte[])dt.Rows[0]["FOTO17_SERIE_POE"];
            MemoryStream mFOTO17_SERIE_POE = new MemoryStream(FOTO17_SERIE_POE);
            byte[] FOTO18_SERIE_AP = (byte[])dt.Rows[0]["FOTO18_SERIE_AP"];
            MemoryStream mFOTO18_SERIE_AP = new MemoryStream(FOTO18_SERIE_AP);
            byte[] FOTO19_SERIE_ANTENA = (byte[])dt.Rows[0]["FOTO19_SERIE_ANTENA"];
            MemoryStream mFOTO19_SERIE_ANTENA = new MemoryStream(FOTO19_SERIE_ANTENA);

            #endregion

            #endregion

            // String usuarioWindows = Environment.UserName;
            String excelGenerado = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";

            File.Copy(rutaPlantilla, excelGenerado, true);

            #region Agregando Valores por hoja en Excel



            #region Caratula 
            ExcelToolsBL.UpdateCell(excelGenerado, "Título", "ESTACION  " + NOMBRE_NODO, 14, "B");
            ExcelToolsBL.UpdateCell(excelGenerado, "Título", CODIGO_NODO, 16, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "Título", TIPO_NODO, 18, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "Título", FRECUENCIA + " Mhz", 22, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "Título", FECHA, 24, "D");
            #endregion

            #region Configuracion y Pruebas
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Pruebas", NOMBRE_NODO, 12, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Pruebas", "ESTACION  " + NOMBRE_NODO, 14, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Pruebas", DIRECCION + "/" + NOMBRE_NODO + "/" + DISTRITO + "/" + PROVINCIA + "/" + DEPARTAMENTO, 15, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Pruebas", LATITUD_S + "/" + LONGITUD_W + "/" + ALTITUD_MSNM, 16, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Pruebas", "Autosoportada Triangular / " + ALTURA_TORRE, 17, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Pruebas", AZIMUT, 21, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Pruebas", DOWN_TILT, 22, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Pruebas", ALTURA_INST, 23, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Pruebas", IP_ADDRESS, 32, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Pruebas", GATEWAY_IP, 34, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Pruebas", FRECUENCIA, 35, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Pruebas", SITE_NAME_SSID, 43, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Pruebas", SITE_NAME_SSID, 47, "E");


            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Pruebas", mCAP1_CONF_RADIO, "",50, 2,873,522);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Pruebas", mCAP2_CONF_QoS, "", 74,2,874,523);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Pruebas", mCAP3_1_CONF_SYSTEM, "",98, 2,391,523);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Pruebas", mCAP3_2_CONF_SYSTEM, "",98,5,484,523);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Pruebas", mCAP4_MONITOR_SYSTEM, "", 122,2,874,469);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Pruebas", mCAP5_1_MON_WIRELESS, "", 146,2,389,421);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Pruebas", mCAP5_2_MON_WIRELESS, "", 146,5,479,412);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Pruebas", mCAP6_TOOLS_WIRELESS, "", 173,2,876,420);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Pruebas", mCAP7_PANTALLA_HOME, "", 200,2,874,455);
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Pruebas", EFICIENCIA_ENLACE, 246, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Pruebas", CAPACIDAD_ENLACE, 247, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Pruebas", NOMBRE_NODO, 249, "E");
            #endregion

            #region Materiales AP

            ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales AP", "ESTACION " + NOMBRE_NODO, 12, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales AP", SERIAL_TERMINAL_ACCESS_POINT, 17, "G");
            ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales AP", SERIAL_ANTENA_CAMBIUM_5GHZ, 20, "G");
            ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales AP", SERIAL_POE_INYECTOR, 25, "G");

            foreach (DataRow dr in dt3.Rows)
            {
                String DESCRIPCION = dr["DESCRIPCION"].ToString();
                String CODIGO = dr["CODIGO"].ToString();
                String UNIDAD = dr["UNIDAD"].ToString();
                String CANTIDAD = dr["CANTIDAD"].ToString();

                int ind = Convert.ToInt32(dt3.Rows.IndexOf(dr));

                ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales AP", Convert.ToString(ind + 1), 32 + ind, "B");
                ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales AP", DESCRIPCION, 32 + ind, "C");
                ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales AP", CODIGO, 32 + ind, "D");
                ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales AP", UNIDAD, 32 + ind, "E");
                ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales AP", CANTIDAD, 32 + ind, "F");
                ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales AP", "S/N", 32 + ind, "G");

            }

            #endregion

            #region SFTP
            ExcelToolsBL.UpdateCell(excelGenerado, "3 SFTP", "ESTACION " + NOMBRE_NODO, 9, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "3 SFTP", AZIMUT + "º", 16, "C");
            ExcelToolsBL.UpdateCell(excelGenerado, "3 SFTP", ALTURA_TORRE, 17, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "3 SFTP", ALTURA_INST, 16, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "3 SFTP", "3", 16, "F");
            ExcelToolsBL.UpdateCell(excelGenerado, "3 SFTP", "2,6", 16, "G");
            ExcelToolsBL.UpdateCell(excelGenerado, "3 SFTP", "8", 16, "H");
            ExcelToolsBL.UpdateCell(excelGenerado, "3 SFTP", "7", 16, "I");
            // ExcelToolsBL.UpdateCell(excelGenerado, "3 SFTP", Convert.ToString(L), 16, "J");
            //  ExcelToolsBL.UpdateCell(excelGenerado, "3 SFTP", Convert.ToString(LT), 16, "L");

            #endregion

            #region Asignaciones

            ExcelToolsBL.UpdateCell(excelGenerado, "4 Asignación", MODELO_PUERTO_NODO, 17, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "4 Asignación", PUERTO_NODO, 17, "F");

            #endregion

            #region Instituciones Atendidas


            foreach (DataRow dr in dt2.Rows)
            {
                String NOMBRE_INST = dr["NOMBRE_INST"].ToString();
                String TIPO_IIBB = dr["TIPO_IIBB"].ToString();
                String LATITUD = dr["LATITUD"].ToString();
                String LONGITUD = dr["LONGITUD"].ToString();

                int ind = Convert.ToInt32(dt2.Rows.IndexOf(dr));

                ExcelToolsBL.UpdateCell(excelGenerado, "6 Instituciones Atendidas", Convert.ToString(ind + 1), 18 + ind, "B");
                ExcelToolsBL.UpdateCell(excelGenerado, "6 Instituciones Atendidas", NOMBRE_INST, 18 + ind, "C");
                ExcelToolsBL.UpdateCell(excelGenerado, "6 Instituciones Atendidas", TIPO_IIBB, 18 + ind, "N");
                ExcelToolsBL.UpdateCell(excelGenerado, "6 Instituciones Atendidas", LATITUD, 18 + ind, "R");
                ExcelToolsBL.UpdateCell(excelGenerado, "6 Instituciones Atendidas", LONGITUD, 18 + ind, "Y");
                ExcelToolsBL.UpdateCell(excelGenerado, "6 Instituciones Atendidas", "SI", 18 + ind, "AE");
            }





            #endregion

            #region Reporte Fotografico

            ExcelToolsBL.UpdateCell(excelGenerado, "7 Reporte fotográfico", "NODO " + NOMBRE_NODO + "  (" + CODIGO_NODO + ")", 12, "B");
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO1_PAN_ESTACION_A, "", 16,3,401,376);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO2_P0S_ANTENA_INST, "", 16, 14,401,376);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO3_ANTENA_ODU_TORRE, "",33,3,402,377);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO4_UGPS, "", 33, 14,403,378);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO5_ENGRASADO_PERNO, "", 50,3,402,414);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO6_SILICONEADO_ETIQUETADO, "", 50,14,401, 414);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO8_RECORRIDO_CABLE_SFTP, "",68,14,403,422);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO8_RECORRIDO_CABLE_SFTP, "", 85,3,401,344);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO10_ATERRAMIENTO_CABLE_SFTP_IN, "",85,14, 404,343);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO11_ETIQUETADO_POE, "",102,3,403,389);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO12_PAN_RACK, "", 102,14,400,388);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO13_ATERRAMIENTO_POE, "",119,3,403,333);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO14_1_EMERGENCIA_POE_ETIQUETA, "", 119, 14, 230,337);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO14_2_EMERGENCIA_POE_ETIQUETA, "", 119, 19, 170,336);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO15_PATCH_CORE_SALIDA_POE, "", 136, 3,403, 346);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO16_PATCH_CORE_SALIDA_SWITCH, "",136,14,403, 343);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO17_SERIE_POE, "", 152, 3,403,342);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO18_SERIE_AP, "", 152,14,403,348);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO19_SERIE_ANTENA, "", 168,3,401,345);
            #endregion


            #region Datos Generales del Nodo

            ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", TIPO_NODO, 9, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", SITE_NAME_SSID, 12, "B");
            ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", NOMBRE_NODO, 12, "G");
            ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", NOMBRE_NODO, 17, "B");
            ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", UBIGEO, 17, "G");
            ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", DEPARTAMENTO, 20, "B");
            ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", PROVINCIA, 20, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", DISTRITO, 20, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", SERIE_ODU, 29, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", SERIE_ANTENA, 30, "I");

            ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", ALTURA_INST + " m", 50, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", DOWN_TILT + " º", 53, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", ALTITUD_MSNM + " m.s.n.m", 54, "I");

            foreach (DataRow dr in dt1.Rows)
            {
                String NODO_A = dr["NODO_A"].ToString();
                String IIBB = dr["IIBB"].ToString();
                String RSS_LOCAL = dr["RSS_LOCAL"].ToString();
                String RSS_REMOTO = dr["RSS_REMOTO"].ToString();
                String TIEMPO_PROM = dr["TIEMPO_PROM"].ToString();
                String CAP_SUBIDA = dr["CAP_SUBIDA"].ToString();
                String CAP_BAJADA = dr["CAP_BAJADA"].ToString();
                String DISTANCIA = dr["DISTANCIA"].ToString();

                int ind = Convert.ToInt32(dt1.Rows.IndexOf(dr));

                ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", NODO_A, 61 + ind, "B");
                ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", IIBB, 61 + ind, "C");
                ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", RSS_LOCAL, 61 + ind, "D");
                ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", RSS_REMOTO, 61 + ind, "E");
                ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", "5", 61 + ind, "F");
                ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", "64QAM 5/6", 61 + ind, "G");
                ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", TIEMPO_PROM, 61 + ind, "H");
                ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", "UL " + CAP_SUBIDA + " /DL " + CAP_BAJADA, 61 + ind, "I");
                ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", DISTANCIA, 61 + ind, "K");
                ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", FRECUENCIA, 61 + ind, "L");

            }

            ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", UBIGEO, 78, "C");
            ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", NOMBRE_NODO, 78, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", DISTRITO, 78, "F");
            ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", PROVINCIA, 78, "H");
            ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", DEPARTAMENTO, 78, "K");
            #endregion

            #endregion

            #region Ruta del Zip
            String rutaNodo = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo;
            if (Directory.Exists(rutaNodo))
            {
                String rutaParcial = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\ACCESO\\PMP\\";
                if (!Directory.Exists(rutaParcial)) Directory.CreateDirectory(rutaParcial);
                String rutaAlterna = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\ACCESO\\PMP\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";
                File.Copy(excelGenerado, rutaAlterna, true);
            }
            #endregion
        }

        public void ActaInstalacionAceptacionProtocoloOmnidireccional(String IdNodo, String IdTarea, String valorCadena1, String rutaPlantilla)
        {
            DBBaseDatos baseDatosDA = new DBBaseDatos();

            baseDatosDA.Configurar();
            baseDatosDA.Conectar();
            DataTable dt = new DataTable();
            DataTable dt1 = new DataTable();
            DataTable dt2 = new DataTable();
            DataTable dt3 = new DataTable();

                try
                {
                    baseDatosDA.CrearComando("USP_R_ACTA_INSTALACION_ACEPTACION_PROTOCOLO_OMNIDIRECCIONAL", CommandType.StoredProcedure);
                    baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                    dt = baseDatosDA.EjecutarConsultaDataTable();

                    baseDatosDA.CrearComando("USP_R_MEDICION_OMNIDIRECCIONAL", CommandType.StoredProcedure);
                    baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                    dt1 = baseDatosDA.EjecutarConsultaDataTable();

                    baseDatosDA.CrearComando("USP_R_IIBB_POR_TAREA", CommandType.StoredProcedure);
                    baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                    dt2 = baseDatosDA.EjecutarConsultaDataTable();

                    baseDatosDA.CrearComando("USP_R_MATERIALES_PROTOCOLO_OMNIDIRECCIONAL", CommandType.StoredProcedure);
                    baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                    dt3 = baseDatosDA.EjecutarConsultaDataTable();

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

            #region Valores


            #region Valores String

            String NOMBRE_NODO = dt.Rows[0]["NOMBRE_NODO"].ToString();
            String CODIGO_NODO = dt.Rows[0]["CODIGO_NODO"].ToString();
            String TIPO_NODO = dt.Rows[0]["TIPO_NODO"].ToString();
            String FRECUENCIA = dt.Rows[0]["FRECUENCIA"].ToString();

            String fechaSQL = dt.Rows[0]["FECHA"].ToString();
            String FECHA = "";
            if (fechaSQL != "")
            {
                DateTime dtFecha = DateTime.Parse(fechaSQL);
                FECHA = dtFecha.ToString("dd/MM/yyyy");
            }
            else { FECHA = ""; }

            String DEPARTAMENTO = dt.Rows[0]["DEPARTAMENTO"].ToString();
            String PROVINCIA = dt.Rows[0]["PROVINCIA"].ToString();
            String DISTRITO = dt.Rows[0]["DISTRITO"].ToString();
            String DIRECCION = dt.Rows[0]["DIRECCION"].ToString();
            String LATITUD_S = dt.Rows[0]["LATITUD_S"].ToString();
            String LONGITUD_W = dt.Rows[0]["LONGITUD_W"].ToString();
            String ALTITUD_MSNM = dt.Rows[0]["ALTITUD_MSNM"].ToString();
            String AZIMUT = dt.Rows[0]["AZIMUT"].ToString();
            String DOWN_TILT = dt.Rows[0]["DOWN_TILT"].ToString();
            String ALTURA_INST = dt.Rows[0]["ALTURA_INST"].ToString();
            String IP_ADDRESS = dt.Rows[0]["IP_ADDRESS"].ToString();
            String GATEWAY_IP = dt.Rows[0]["GATEWAY_IP"].ToString();
            String CAPACIDAD_ENLACE = dt.Rows[0]["CAPACIDAD_ENLACE"].ToString();
            String EFICIENCIA_ENLACE = dt.Rows[0]["EFICIENCIA_ENLACE"].ToString();
            String SITE_NAME_SSID = dt.Rows[0]["SITE_NAME_SSID"].ToString();

            //String DISTANCIA_B = dt.Rows[0]["DISTANCIA_B"].ToString();
            //String DISTANCIA_C = dt.Rows[0]["DISTANCIA_C"].ToString();
            //String DISTANCIA_D = dt.Rows[0]["DISTANCIA_D"].ToString();
            //String DISTANCIA_E = dt.Rows[0]["DISTANCIA_E"].ToString();
            String ALTURA_TORRE = dt.Rows[0]["ALTURA_TORRE"].ToString();

            String SERIE_ODU = dt.Rows[0]["#SERIE_ODU"].ToString();
            String SERIE_ANTENA = dt.Rows[0]["#SERIE_ANTENA"].ToString();

            String SERIAL_TERMINAL_ACCESS_POINT = dt.Rows[0]["SERIAL_TERMINAL_ACCESS_POINT"].ToString();
            String SERIAL_ANTENA_CAMBIUM_5GHZ = dt.Rows[0]["SERIAL_ANTENA_CAMBIUM_5GHZ"].ToString();
            String SERIAL_POE_INYECTOR = dt.Rows[0]["SERIAL_POE_INYECTOR"].ToString();

            String UBIGEO = dt.Rows[0]["UBIGEO"].ToString();

            String MODELO_PUERTO_NODO = dt.Rows[0]["MODELO_PUERTO_NODO"].ToString();
            String PUERTO_NODO = dt.Rows[0]["PUERTO_NODO"].ToString();

            /*   int a = Convert.ToInt32(ALTURA_INST);
         int b = Convert.ToInt32(DISTANCIA_B);
         int c = Convert.ToInt32(DISTANCIA_C);
         int d = Convert.ToInt32(DISTANCIA_D);
         int e = Convert.ToInt32(DISTANCIA_E);

         int L = a + b - c + d + e;

         int LT = L + 3;  */

            #endregion

            #region Valores Binarios

            byte[] CAP1_CONF_RADIO = (byte[])dt.Rows[0]["CAP1_CONF_RADIO"];
            MemoryStream mCAP1_CONF_RADIO = new MemoryStream(CAP1_CONF_RADIO);
            byte[] CAP2_CONF_QoS = (byte[])dt.Rows[0]["CAP2_CONF_QoS"];
            MemoryStream mCAP2_CONF_QoS = new MemoryStream(CAP2_CONF_QoS);
            byte[] CAP3_1_CONF_SYSTEM = (byte[])dt.Rows[0]["CAP3_1_CONF_SYSTEM"];
            MemoryStream mCAP3_1_CONF_SYSTEM = new MemoryStream(CAP3_1_CONF_SYSTEM);
            byte[] CAP3_2_CONF_SYSTEM = (byte[])dt.Rows[0]["CAP3_2_CONF_SYSTEM"];
            MemoryStream mCAP3_2_CONF_SYSTEM = new MemoryStream(CAP3_2_CONF_SYSTEM);
            byte[] CAP4_MONITOR_SYSTEM = (byte[])dt.Rows[0]["CAP4_MONITOR_SYSTEM"];
            MemoryStream mCAP4_MONITOR_SYSTEM = new MemoryStream(CAP4_MONITOR_SYSTEM);
            byte[] CAP5_1_MON_WIRELESS = (byte[])dt.Rows[0]["CAP5_1_MON_WIRELESS"];
            MemoryStream mCAP5_1_MON_WIRELESS = new MemoryStream(CAP5_1_MON_WIRELESS);
            byte[] CAP5_2_MON_WIRELESS = (byte[])dt.Rows[0]["CAP5_2_MON_WIRELESS"];
            MemoryStream mCAP5_2_MON_WIRELESS = new MemoryStream(CAP5_2_MON_WIRELESS);
            byte[] CAP6_TOOLS_WIRELESS = (byte[])dt.Rows[0]["CAP6_TOOLS_WIRELESS"];
            MemoryStream mCAP6_TOOLS_WIRELESS = new MemoryStream(CAP6_TOOLS_WIRELESS);
            byte[] CAP7_PANTALLA_HOME = (byte[])dt.Rows[0]["CAP7_PANTALLA_HOME"];
            MemoryStream mCAP7_PANTALLA_HOME = new MemoryStream(CAP7_PANTALLA_HOME);

            byte[] FOTO1_PAN_ESTACION_A = (byte[])dt.Rows[0]["FOTO1_PAN_ESTACION_A"];
            MemoryStream mFOTO1_PAN_ESTACION_A = new MemoryStream(FOTO1_PAN_ESTACION_A);
            byte[] FOTO2_P0S_ANTENA_INST = (byte[])dt.Rows[0]["FOTO2_P0S_ANTENA_INST"];
            MemoryStream mFOTO2_P0S_ANTENA_INST = new MemoryStream(FOTO2_P0S_ANTENA_INST);
            byte[] FOTO3_ANTENA_ODU_TORRE = (byte[])dt.Rows[0]["FOTO3_ANTENA_ODU_TORRE"];
            MemoryStream mFOTO3_ANTENA_ODU_TORRE = new MemoryStream(FOTO3_ANTENA_ODU_TORRE);
            byte[] FOTO4_UGPS = (byte[])dt.Rows[0]["FOTO4_UGPS"];
            MemoryStream mFOTO4_UGPS = new MemoryStream(FOTO4_UGPS);
            byte[] FOTO5_ENGRASADO_PERNO = (byte[])dt.Rows[0]["FOTO5_ENGRASADO_PERNO"];
            MemoryStream mFOTO5_ENGRASADO_PERNO = new MemoryStream(FOTO5_ENGRASADO_PERNO);
            byte[] FOTO6_SILICONEADO_ETIQUETADO = (byte[])dt.Rows[0]["FOTO6_SILICONEADO_ETIQUETADO"];
            MemoryStream mFOTO6_SILICONEADO_ETIQUETADO = new MemoryStream(FOTO6_SILICONEADO_ETIQUETADO);
            byte[] FOTO8_RECORRIDO_CABLE_SFTP = (byte[])dt.Rows[0]["FOTO8_RECORRIDO_CABLE_SFTP"];
            MemoryStream mFOTO8_RECORRIDO_CABLE_SFTP = new MemoryStream(FOTO8_RECORRIDO_CABLE_SFTP);
            byte[] FOTO9_ATERRAMIENTO_CABLE_SFTP_OUT = (byte[])dt.Rows[0]["FOTO9_ATERRAMIENTO_CABLE_SFTP_OUT"];
            MemoryStream mFOTO9_ATERRAMIENTO_CABLE_SFTP_OUT = new MemoryStream(FOTO9_ATERRAMIENTO_CABLE_SFTP_OUT);
            byte[] FOTO10_ATERRAMIENTO_CABLE_SFTP_IN = (byte[])dt.Rows[0]["FOTO10_ATERRAMIENTO_CABLE_SFTP_IN"];
            MemoryStream mFOTO10_ATERRAMIENTO_CABLE_SFTP_IN = new MemoryStream(FOTO10_ATERRAMIENTO_CABLE_SFTP_IN);
            byte[] FOTO11_ETIQUETADO_POE = (byte[])dt.Rows[0]["FOTO11_ETIQUETADO_POE"];
            MemoryStream mFOTO11_ETIQUETADO_POE = new MemoryStream(FOTO11_ETIQUETADO_POE);
            byte[] FOTO12_PAN_RACK = (byte[])dt.Rows[0]["FOTO12_PAN_RACK"];
            MemoryStream mFOTO12_PAN_RACK = new MemoryStream(FOTO12_PAN_RACK);
            byte[] FOTO13_ATERRAMIENTO_POE = (byte[])dt.Rows[0]["FOTO13_ATERRAMIENTO_POE"];
            MemoryStream mFOTO13_ATERRAMIENTO_POE = new MemoryStream(FOTO13_ATERRAMIENTO_POE);
            byte[] FOTO14_1_EMERGENCIA_POE_ETIQUETA = (byte[])dt.Rows[0]["FOTO14_1_EMERGENCIA_POE_ETIQUETA"];
            MemoryStream mFOTO14_1_EMERGENCIA_POE_ETIQUETA = new MemoryStream(FOTO14_1_EMERGENCIA_POE_ETIQUETA);
            byte[] FOTO14_2_EMERGENCIA_POE_ETIQUETA = (byte[])dt.Rows[0]["FOTO14_2_EMERGENCIA_POE_ETIQUETA"];
            MemoryStream mFOTO14_2_EMERGENCIA_POE_ETIQUETA = new MemoryStream(FOTO14_2_EMERGENCIA_POE_ETIQUETA);
            byte[] FOTO15_PATCH_CORE_SALIDA_POE = (byte[])dt.Rows[0]["FOTO15_PATCH_CORE_SALIDA_POE"];
            MemoryStream mFOTO15_PATCH_CORE_SALIDA_POE = new MemoryStream(FOTO15_PATCH_CORE_SALIDA_POE);
            byte[] FOTO16_PATCH_CORE_SALIDA_SWITCH = (byte[])dt.Rows[0]["FOTO16_PATCH_CORE_SALIDA_SWITCH"];
            MemoryStream mFOTO16_PATCH_CORE_SALIDA_SWITCH = new MemoryStream(FOTO16_PATCH_CORE_SALIDA_SWITCH);
            byte[] FOTO17_SERIE_POE = (byte[])dt.Rows[0]["FOTO17_SERIE_POE"];
            MemoryStream mFOTO17_SERIE_POE = new MemoryStream(FOTO17_SERIE_POE);
            byte[] FOTO18_SERIE_AP = (byte[])dt.Rows[0]["FOTO18_SERIE_AP"];
            MemoryStream mFOTO18_SERIE_AP = new MemoryStream(FOTO18_SERIE_AP);
            byte[] FOTO19_SERIE_ANTENA = (byte[])dt.Rows[0]["FOTO19_SERIE_ANTENA"];
            MemoryStream mFOTO19_SERIE_ANTENA = new MemoryStream(FOTO19_SERIE_ANTENA);

            #endregion

            #endregion

            // String usuarioWindows = Environment.UserName;
            String excelGenerado = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";

            File.Copy(rutaPlantilla, excelGenerado, true);

            #region Agregando Valores por hoja en Excel



            #region Caratula 
            ExcelToolsBL.UpdateCell(excelGenerado, "Título", "ESTACION  " + NOMBRE_NODO, 14, "B");
            ExcelToolsBL.UpdateCell(excelGenerado, "Título", CODIGO_NODO, 16, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "Título", TIPO_NODO, 18, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "Título", FRECUENCIA + " Mhz", 22, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "Título", FECHA, 24, "D");
            #endregion

            #region Configuracion y Pruebas
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Pruebas", NOMBRE_NODO, 12, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Pruebas", "ESTACION  " + NOMBRE_NODO, 14, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Pruebas", DIRECCION + "/" + NOMBRE_NODO + "/" + DISTRITO + "/" + PROVINCIA + "/" + DEPARTAMENTO, 15, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Pruebas", LATITUD_S + "/" + LONGITUD_W + "/" + ALTITUD_MSNM, 16, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Pruebas", "Autosoportada Triangular / " + ALTURA_TORRE, 17, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Pruebas", AZIMUT, 21, "H");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Pruebas", DOWN_TILT, 22, "H");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Pruebas", ALTURA_INST, 23, "H");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Pruebas", IP_ADDRESS, 32, "H");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Pruebas", GATEWAY_IP, 34, "H");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Pruebas", FRECUENCIA, 35, "H");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Pruebas", SITE_NAME_SSID, 43, "H");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Pruebas", SITE_NAME_SSID, 47, "H");


            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Pruebas", mCAP1_CONF_RADIO, "", 50, 2, 851, 514);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Pruebas", mCAP2_CONF_QoS, "", 74, 2,848, 513);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Pruebas", mCAP3_1_CONF_SYSTEM, "", 98, 2, 487, 515);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Pruebas", mCAP3_2_CONF_SYSTEM, "", 98, 6, 363, 513);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Pruebas", mCAP4_MONITOR_SYSTEM, "", 122, 2, 852, 457);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Pruebas", mCAP5_1_MON_WIRELESS, "", 146, 2, 369, 408);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Pruebas", mCAP5_2_MON_WIRELESS, "", 146, 5, 478, 407);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Pruebas", mCAP6_TOOLS_WIRELESS, "", 173, 2, 848, 405);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Pruebas", mCAP7_PANTALLA_HOME, "", 200, 2, 849, 443);
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Pruebas", EFICIENCIA_ENLACE, 246, "H");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Pruebas", CAPACIDAD_ENLACE, 247, "H");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Pruebas", NOMBRE_NODO, 249, "H");
            #endregion

            #region Materiales AP

            ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales AP", "ESTACION " + NOMBRE_NODO, 12, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales AP", SERIAL_TERMINAL_ACCESS_POINT, 17, "G");
            ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales AP", SERIAL_ANTENA_CAMBIUM_5GHZ, 20, "G");
            ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales AP", SERIAL_POE_INYECTOR, 25, "G");

            foreach (DataRow dr in dt3.Rows)
            {
                String DESCRIPCION = dr["DESCRIPCION"].ToString();
                String CODIGO = dr["CODIGO"].ToString();
                String UNIDAD = dr["UNIDAD"].ToString();
                String CANTIDAD = dr["CANTIDAD"].ToString();

                int ind = Convert.ToInt32(dt3.Rows.IndexOf(dr));

                ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales AP", Convert.ToString(ind + 1), 32 + ind, "B");
                ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales AP", DESCRIPCION, 32 + ind, "C");
                ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales AP", CODIGO, 32 + ind, "D");
                ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales AP", UNIDAD, 32 + ind, "E");
                ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales AP", CANTIDAD, 32 + ind, "F");
                ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales AP", "S/N", 32 + ind, "G");

            }

            #endregion

            #region SFTP
            ExcelToolsBL.UpdateCell(excelGenerado, "3 SFTP", "ESTACION " + NOMBRE_NODO, 9, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "3 SFTP", AZIMUT + "º", 16, "C");
            ExcelToolsBL.UpdateCell(excelGenerado, "3 SFTP", ALTURA_TORRE, 16, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "3 SFTP", ALTURA_INST, 16, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "3 SFTP", "3", 16, "F");
            ExcelToolsBL.UpdateCell(excelGenerado, "3 SFTP", "2,6", 16, "G");
            ExcelToolsBL.UpdateCell(excelGenerado, "3 SFTP", "8", 16, "H");
            ExcelToolsBL.UpdateCell(excelGenerado, "3 SFTP", "7", 16, "I");
            // ExcelToolsBL.UpdateCell(excelGenerado, "3 SFTP", Convert.ToString(L), 16, "J");
            //  ExcelToolsBL.UpdateCell(excelGenerado, "3 SFTP", Convert.ToString(LT), 16, "L");

            #endregion

            #region Asignaciones

            ExcelToolsBL.UpdateCell(excelGenerado, "4 Asignación", MODELO_PUERTO_NODO, 17, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "4 Asignación", PUERTO_NODO, 17, "F");

            #endregion

            #region Instituciones Atendidas


            foreach (DataRow dr in dt2.Rows)
            {
                String NOMBRE_INST = dr["NOMBRE_INST"].ToString();
                String TIPO_IIBB = dr["TIPO_IIBB"].ToString();
                String LATITUD = dr["LATITUD"].ToString();
                String LONGITUD = dr["LONGITUD"].ToString();

                int ind = Convert.ToInt32(dt2.Rows.IndexOf(dr));

                ExcelToolsBL.UpdateCell(excelGenerado, "6 Instituciones Atendidas", Convert.ToString(ind + 1), 18 + ind, "B");
                ExcelToolsBL.UpdateCell(excelGenerado, "6 Instituciones Atendidas", NOMBRE_INST, 18 + ind, "C");
                ExcelToolsBL.UpdateCell(excelGenerado, "6 Instituciones Atendidas", TIPO_IIBB, 18 + ind, "N");
                ExcelToolsBL.UpdateCell(excelGenerado, "6 Instituciones Atendidas", LATITUD, 18 + ind, "R");
                ExcelToolsBL.UpdateCell(excelGenerado, "6 Instituciones Atendidas", LONGITUD, 18 + ind, "Y");
                ExcelToolsBL.UpdateCell(excelGenerado, "6 Instituciones Atendidas", "SI", 18 + ind, "AE");
            }





            #endregion

            #region Reporte Fotografico

            ExcelToolsBL.UpdateCell(excelGenerado, "7 Reporte fotográfico", "NODO " + NOMBRE_NODO + "  (" + CODIGO_NODO + ")", 12, "B");
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO1_PAN_ESTACION_A, "", 16, 3, 402, 380);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO2_P0S_ANTENA_INST, "", 16, 14, 403, 382);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO3_ANTENA_ODU_TORRE, "", 33, 3, 402, 384);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO4_UGPS, "", 33, 14, 402, 383);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO5_ENGRASADO_PERNO, "", 50, 3, 404, 419);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO6_SILICONEADO_ETIQUETADO, "", 50, 14, 403, 419);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO8_RECORRIDO_CABLE_SFTP, "", 68, 14, 403, 429);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO8_RECORRIDO_CABLE_SFTP, "", 85, 3, 403, 347);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO10_ATERRAMIENTO_CABLE_SFTP_IN, "", 86, 14, 404, 349);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO11_ETIQUETADO_POE, "", 102, 3, 404, 389);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO12_PAN_RACK, "", 102, 14, 405, 389);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO13_ATERRAMIENTO_POE, "", 119, 3, 404, 338);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO14_1_EMERGENCIA_POE_ETIQUETA, "", 119, 14, 223, 340);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO14_2_EMERGENCIA_POE_ETIQUETA, "", 119, 19, 180, 338);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO15_PATCH_CORE_SALIDA_POE, "", 136, 3, 405, 348);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO16_PATCH_CORE_SALIDA_SWITCH, "", 136, 14, 404, 349);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO17_SERIE_POE, "", 152, 3, 404, 348);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO18_SERIE_AP, "", 152, 14, 404, 348);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte fotográfico", mFOTO19_SERIE_ANTENA, "", 168, 3, 404, 349);
            #endregion


            #region Datos Generales del Nodo

            ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", TIPO_NODO, 9, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", SITE_NAME_SSID, 12, "B");
            ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", NOMBRE_NODO, 12, "G");
            ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", NOMBRE_NODO, 17, "B");
            ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", UBIGEO, 17, "G");
            ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", DEPARTAMENTO, 20, "B");
            ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", PROVINCIA, 20, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", DISTRITO, 20, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", SERIE_ODU, 29, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", SERIE_ANTENA, 30, "I");

            ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", ALTURA_INST + " m", 50, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", DOWN_TILT + " º", 53, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", ALTITUD_MSNM + " m.s.n.m", 54, "I");

            foreach (DataRow dr in dt1.Rows)
            {
                String NODO_A = dr["NODO_A"].ToString();
                String IIBB = dr["IIBB"].ToString();
                String RSS_LOCAL = dr["RSS_LOCAL"].ToString();
                String RSS_REMOTO = dr["RSS_REMOTO"].ToString();
                String TIEMPO_PROM = dr["TIEMPO_PROM"].ToString();
                String CAP_SUBIDA = dr["CAP_SUBIDA"].ToString();
                String CAP_BAJADA = dr["CAP_BAJADA"].ToString();
                String DISTANCIA = dr["DISTANCIA"].ToString();

                int ind = Convert.ToInt32(dt1.Rows.IndexOf(dr));

                ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", NODO_A, 61 + ind, "B");
                ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", IIBB, 61 + ind, "C");
                ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", RSS_LOCAL, 61 + ind, "D");
                ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", RSS_REMOTO, 61 + ind, "E");
                ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", "5", 61 + ind, "F");
                ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", "64QAM 5/6", 61 + ind, "G");
                ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", TIEMPO_PROM, 61 + ind, "H");
                ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", "UL " + CAP_SUBIDA + " /DL " + CAP_BAJADA, 61 + ind, "I");
                ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", DISTANCIA, 61 + ind, "K");
                ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", FRECUENCIA, 61 + ind, "L");

            }

            ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", UBIGEO, 78, "C");
            ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", NOMBRE_NODO, 78, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", DISTRITO, 78, "F");
            ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", PROVINCIA, 78, "H");
            ExcelToolsBL.UpdateCell(excelGenerado, "8 Datos generales del nodo", DEPARTAMENTO, 78, "K");
            #endregion

            #endregion

            #region Ruta del Zip
            String rutaNodo = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo;
            if (Directory.Exists(rutaNodo))
            {
                String rutaParcial = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\ACCESO\\PMP\\";
                if (!Directory.Exists(rutaParcial)) Directory.CreateDirectory(rutaParcial);
                String rutaAlterna = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\ACCESO\\PMP\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";
                File.Copy(excelGenerado, rutaAlterna, true);
            }
            #endregion
        }

        public void ActaInstalacionAceptacionProtocoloIIBB_A(String IdNodo, String IdTarea, String valorCadena1, String rutaPlantilla)
        {
            DBBaseDatos baseDatosDA = new DBBaseDatos();

            baseDatosDA.Configurar();
            baseDatosDA.Conectar();

            DataTable dt = new DataTable();
            DataTable dt1 = new DataTable();
            DataTable dt2 = new DataTable();
            DataTable dt3 = new DataTable();

            try
            {
                baseDatosDA.CrearComando("USP_R_ACTA_INSTALACION_ACEPTACION_PROTOCOLO_IIBB_A", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                 dt = baseDatosDA.EjecutarConsultaDataTable();

                baseDatosDA.CrearComando("USP_R_MEDICION_ENLACE_IIBB_A", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                 dt1 = baseDatosDA.EjecutarConsultaDataTable();

                baseDatosDA.CrearComando("USP_R_EQUIPAMIENTOS_IIBB_A", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                 dt2 = baseDatosDA.EjecutarConsultaDataTable();

                baseDatosDA.CrearComando("USP_R_MATERIALES_IIBB_A", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                 dt3 = baseDatosDA.EjecutarConsultaDataTable();


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


            #region Valores

            #region Valores String

            String FRECUENCIA = dt.Rows[0]["FRECUENCIA"].ToString();
            String CODIGO_IIBB = dt.Rows[0]["CODIGO_IIBB"].ToString();
            String TIPO_INSTITUCION = dt.Rows[0]["TIPO_INSTITUCION"].ToString();
            String NOMBRE_IIBB = dt.Rows[0]["NOMBRE_IIBB"].ToString();

            String fechaSQL = dt.Rows[0]["FECHA"].ToString();
            String FECHA = "";
            if (fechaSQL != "")
            {
                DateTime dtFecha = DateTime.Parse(fechaSQL);
                FECHA = dtFecha.ToString("dd/MM/yyyy");
            }
            else { FECHA = ""; }

            String DEPARTAMENTO = dt.Rows[0]["DEPARTAMENTO"].ToString();
            String PROVINCIA = dt.Rows[0]["PROVINCIA"].ToString();
            String DISTRITO = dt.Rows[0]["DISTRITO"].ToString();
            String NOMBRE_NODO = dt.Rows[0]["NOMBRE_NODO"].ToString();
            String UBIGEO = dt.Rows[0]["UBIGEO"].ToString();
            String LATITUD = dt.Rows[0]["LATITUD"].ToString();
            String LONGITUD = dt.Rows[0]["LONGITUD"].ToString();
            String ALTITUDmsnm = dt.Rows[0]["ALTITUDmsnm"].ToString();
            String AZIMUT = dt.Rows[0]["AZIMUT"].ToString();


            String DIRECCION_NODO = dt.Rows[0]["DIRECCION_NODO"].ToString();

            String ODU_CPE = dt.Rows[0]["ODU_CPE"].ToString();
            String ACCESS_POINT_INDOOR = dt.Rows[0]["ACCESS_POINT_INDOOR"].ToString();
            String SWITCH_COMUNICACIONES = dt.Rows[0]["SWITCH_COMUNICACIONES"].ToString();
            String ROUTER = dt.Rows[0]["ROUTER"].ToString();
            String EQUIPO_COMPUTO1 = dt.Rows[0]["EQUIPO_COMPUTO1"].ToString();

            String IMPRESORA_MULTIFUNCIONAL = dt.Rows[0]["IMPRESORA_MULTIFUNCIONAL"].ToString();
            String UPS = dt.Rows[0]["UPS"].ToString();
            String REFERENCIA_UBICACION_IIBB = dt.Rows[0]["REFERENCIA_UBICACION_IIBB"].ToString();
            String TIPO_MASTIL = dt.Rows[0]["TIPO_MASTIL"].ToString();
            String ALTURA_MASTIL = dt.Rows[0]["ALTURA_MASTIL"].ToString();
            String DISPONIBILIDAD_HORAS = dt.Rows[0]["DISPONIBILIDAD_HORAS"].ToString();

            String POTENCIA_TRANSMISION = dt.Rows[0]["POTENCIA_TRANSMISION"].ToString();
         //   String ANCHO_BANDA_CANAL = dt.Rows[0]["ANCHO_BANDA_CANAL"].ToString();  (VALOR FIJO 20, SEGUN ARCHIVO SOFTWARE(1))
            String ELEVACION = dt.Rows[0]["ELEVACION"].ToString();
            String CONECTIVIDAD_GILAT = dt.Rows[0]["CONECTIVIDAD_GILAT"].ToString();
            String CONECTIVIDAD_NODO_TERMINAL = dt.Rows[0]["CONECTIVIDAD_NODO_TERMINAL"].ToString();
            String CONECTIVIDAD_NODO_DISTRITAL = dt.Rows[0]["CONECTIVIDAD_NODO_DISTRITAL"].ToString();
            String CONECTIVIDAD_NOC = dt.Rows[0]["CONECTIVIDAD_NOC"].ToString();
            String NOMBRES_APELLIDOS_ENCARGADO = dt.Rows[0]["NOMBRES_APELLIDOS_ENCARGADO"].ToString();
            String DOC_IDENTIDAD_ENCARGADO = dt.Rows[0]["DOC_IDENTIDAD_ENCARGADO"].ToString();
            String CELULAR_CONTACTO_ENCARGADO = dt.Rows[0]["CELULAR_CONTACTO_ENCARGADO"].ToString();
            String EMAIL_ENCARGADO_IIBB = dt.Rows[0]["EMAIL_ENCARGADO_IIBB"].ToString();
            String NOMBRES_APELLIDOS_REPRESENTANTE = dt.Rows[0]["NOMBRES_APELLIDOS_REPRESENTANTE"].ToString();
            String DOC_IDENTIDAD_REPRESENTANTE = dt.Rows[0]["DOC_IDENTIDAD_REPRESENTANTE"].ToString();
            String CELULAR_CONTACTO_REPRESENTANTE = dt.Rows[0]["CELULAR_CONTACTO_REPRESENTANTE"].ToString();
            String CARGO_REPRESENTANTE_IIBB = dt.Rows[0]["CARGO_REPRESENTANTE_IIBB"].ToString();
            String EMAIL_REPRESENTANTE_IIBB = dt.Rows[0]["EMAIL_REPRESENTANTE_IIBB"].ToString();

            //TODOS ESTOS VALORES NO APLICAN SEGUN ARCHIVO SOFTWARE (1)
            //String MSPT_MEDIDA1_VALORMEDIO = dt.Rows[0]["MSPT_MEDIDA1_VALORMEDIO"].ToString();
            //String MSPT_MEDIDA2_VALORMEDIO = dt.Rows[0]["MSPT_MEDIDA2_VALORMEDIO"].ToString();
            //String MSPT_MEDIDA3_VALORMEDIO = dt.Rows[0]["MSPT_MEDIDA3_VALORMEDIO"].ToString();
            String MSPTP_MEDIDA1_VALORMEDIO = dt.Rows[0]["MSPTP_MEDIDA1_VALORMEDIO"].ToString();
            String MSPTP_MEDIDA2_VALORMEDIO = dt.Rows[0]["MSPTP_MEDIDA2_VALORMEDIO"].ToString();
            String MSPTP_MEDIDA3_VALORMEDIO = dt.Rows[0]["MSPTP_MEDIDA3_VALORMEDIO"].ToString();

            String CODIGO_NODO = dt.Rows[0]["CODIGO_NODO"].ToString();
            String IP_IIBB = dt.Rows[0]["IP_IIBB"].ToString();



            #endregion

            #region Valores byte

            byte[] PANT_CONF_ACCESS_POINt = (byte[])dt.Rows[0]["PANT_CONF_ACCESS_POINt"];
            MemoryStream mPANT_CONF_ACCESS_POINt = new MemoryStream(PANT_CONF_ACCESS_POINt);
            byte[] PANT_CONF_ROUTER = (byte[])dt.Rows[0]["PANT_CONF_ROUTER"];
            MemoryStream mPANT_CONF_ROUTER = new MemoryStream(PANT_CONF_ROUTER);
            byte[] PANT_CONF_SWITCH01 = (byte[])dt.Rows[0]["PANT_CONF_SWITCH01"];
            MemoryStream mPANT_CONF_SWITCH01 = new MemoryStream(PANT_CONF_SWITCH01);
            byte[] PANT_CONF_SWITCH02 = (byte[])dt.Rows[0]["PANT_CONF_SWITCH02"];
            MemoryStream mPANT_CONF_SWITCH02 = new MemoryStream(PANT_CONF_SWITCH02);
            byte[] PANT_CONF_UPS = (byte[])dt.Rows[0]["PANT_CONF_UPS"];
            MemoryStream mPANT_CONF_UPS = new MemoryStream(PANT_CONF_UPS);
            byte[] PANT_CONF_ALLINONE01 = (byte[])dt.Rows[0]["PANT_CONF_ALLINONE01"];
            MemoryStream mPANT_CONF_ALLINONE01 = new MemoryStream(PANT_CONF_ALLINONE01);
            byte[] PANT_CONF_ALLINONE02 = (byte[])dt.Rows[0]["PANT_CONF_ALLINONE02"];
            MemoryStream mPANT_CONF_ALLINONE02 = new MemoryStream(PANT_CONF_ALLINONE02);
            byte[] PANT_CONF_IMPRESORA = (byte[])dt.Rows[0]["PANT_CONF_IMPRESORA"];
            MemoryStream mPANT_CONF_IMPRESORA = new MemoryStream(PANT_CONF_IMPRESORA);
            byte[] FOTO1_PAN_LOCALIDAD = (byte[])dt.Rows[0]["FOTO1_PAN_LOCALIDAD"];
            MemoryStream mFOTO1_PAN_LOCALIDAD = new MemoryStream(FOTO1_PAN_LOCALIDAD);
            byte[] FOTO2_FACHADA_INSTITUCION = (byte[])dt.Rows[0]["FOTO2_FACHADA_INSTITUCION"];
            MemoryStream mFOTO2_FACHADA_INSTITUCION = new MemoryStream(FOTO2_FACHADA_INSTITUCION);
            byte[] FOTO3_1_IMPRESORA = (byte[])dt.Rows[0]["FOTO3_1_IMPRESORA"];
            MemoryStream mFOTO3_1_IMPRESORA = new MemoryStream(FOTO3_1_IMPRESORA);
            byte[] FOTO3_2_SWITCH = (byte[])dt.Rows[0]["FOTO3_2_SWITCH"];
            MemoryStream mFOTO3_2_SWITCH = new MemoryStream(FOTO3_2_SWITCH);
            byte[] FOTO3_3_ROUTER = (byte[])dt.Rows[0]["FOTO3_3_ROUTER"];
            MemoryStream mFOTO3_3_ROUTER = new MemoryStream(FOTO3_3_ROUTER);
            byte[] FOTO3_4_PC_ENCENDIDAS = (byte[])dt.Rows[0]["FOTO3_4_PC_ENCENDIDAS"];
            MemoryStream mFOTO3_4_PC_ENCENDIDAS = new MemoryStream(FOTO3_4_PC_ENCENDIDAS);
            byte[] FOTO3_5_PC_UPS = (byte[])dt.Rows[0]["FOTO3_5_PC_UPS"];
            MemoryStream mFOTO3_5_PC_UPS = new MemoryStream(FOTO3_5_PC_UPS);
            byte[] FOTO3_6_ACCESS_POINT = (byte[])dt.Rows[0]["FOTO3_6_ACCESS_POINT"];
            MemoryStream mFOTO3_6_ACCESS_POINT = new MemoryStream(FOTO3_6_ACCESS_POINT);
            byte[] FOTO4_1_ODU_CPE = (byte[])dt.Rows[0]["FOTO4_1_ODU_CPE"];
            MemoryStream mFOTO4_1_ODU_CPE = new MemoryStream(FOTO4_1_ODU_CPE);
            byte[] FOTO4_2_MASTIL = (byte[])dt.Rows[0]["FOTO4_2_MASTIL"];
            MemoryStream mFOTO4_2_MASTIL = new MemoryStream(FOTO4_2_MASTIL);
            byte[] FOTO4_3_PAN_ANT_INSTAL_MASTIL = (byte[])dt.Rows[0]["FOTO4_3_PAN_ANT_INSTAL_MASTIL"];
            MemoryStream mFOTO4_3_PAN_ANT_INSTAL_MASTIL = new MemoryStream(FOTO4_3_PAN_ANT_INSTAL_MASTIL);
            byte[] FOTO4_4_RECORRIDO_SFTP_CATSE = (byte[])dt.Rows[0]["FOTO4_4_RECORRIDO_SFTP_CATSE"];
            MemoryStream mFOTO4_4_RECORRIDO_SFTP_CATSE = new MemoryStream(FOTO4_4_RECORRIDO_SFTP_CATSE);
            byte[] FOTO4_5_INGRESO_SFTP = (byte[])dt.Rows[0]["FOTO4_5_INGRESO_SFTP"];
            MemoryStream mFOTO4_5_INGRESO_SFTP = new MemoryStream(FOTO4_5_INGRESO_SFTP);
            byte[] FOTO4_6_RECORRIDO_SFTP_CANALETA = (byte[])dt.Rows[0]["FOTO4_6_RECORRIDO_SFTP_CANALETA"];
            MemoryStream mFOTO4_6_RECORRIDO_SFTP_CANALETA = new MemoryStream(FOTO4_6_RECORRIDO_SFTP_CANALETA);
            byte[] FOTO4_7_POE = (byte[])dt.Rows[0]["FOTO4_7_POE"];
            MemoryStream mFOTO4_7_POE = new MemoryStream(FOTO4_7_POE);
            byte[] FOTO4_8_PATCH_POE_ROUTER = (byte[])dt.Rows[0]["FOTO4_8_PATCH_POE_ROUTER"];
            MemoryStream mFOTO4_8_PATCH_POE_ROUTER = new MemoryStream(FOTO4_8_PATCH_POE_ROUTER);
            byte[] FOTO5_1_TABLERO_GENERAL_SECUNDARIO = (byte[])dt.Rows[0]["FOTO5_1_TABLERO_GENERAL_SECUNDARIO"];
            MemoryStream mFOTO5_1_TABLERO_GENERAL_SECUNDARIO = new MemoryStream(FOTO5_1_TABLERO_GENERAL_SECUNDARIO);
            byte[] FOTO5_2_INSTALACION_BREAKER = (byte[])dt.Rows[0]["FOTO5_2_INSTALACION_BREAKER"];
            MemoryStream mFOTO5_2_INSTALACION_BREAKER = new MemoryStream(FOTO5_2_INSTALACION_BREAKER);
            byte[] FOTO5_3_CABLE_CONEXION_ELECTRICA = (byte[])dt.Rows[0]["FOTO5_3_CABLE_CONEXION_ELECTRICA"];
            MemoryStream mFOTO5_3_CABLE_CONEXION_ELECTRICA = new MemoryStream(FOTO5_3_CABLE_CONEXION_ELECTRICA);
            byte[] FOTO5_4_TOMAS_ENERGIA = (byte[])dt.Rows[0]["FOTO5_4_TOMAS_ENERGIA"];
            MemoryStream mFOTO5_4_TOMAS_ENERGIA = new MemoryStream(FOTO5_4_TOMAS_ENERGIA);
            byte[] FOTO5_5_FOTO_INTERNA_INST_BREAKER = (byte[])dt.Rows[0]["FOTO5_5_FOTO_INTERNA_INST_BREAKER"];
            MemoryStream mFOTO5_5_FOTO_INTERNA_INST_BREAKER = new MemoryStream(FOTO5_5_FOTO_INTERNA_INST_BREAKER);
            byte[] FOTO6_1_DNI_DJREPRESENTANTE_ABONADO = (byte[])dt.Rows[0]["FOTO6_1_DNI_DJREPRESENTANTE_ABONADO"];
            MemoryStream mFOTO6_1_DNI_DJREPRESENTANTE_ABONADO = new MemoryStream(FOTO6_1_DNI_DJREPRESENTANTE_ABONADO);
            byte[] FOTO6_2_DNI_DJREPRESENTANTE_ABONADO = (byte[])dt.Rows[0]["FOTO6_2_DNI_DJREPRESENTANTE_ABONADO"];
            MemoryStream mFOTO6_2_DNI_DJREPRESENTANTE_ABONADO = new MemoryStream(FOTO6_2_DNI_DJREPRESENTANTE_ABONADO);
            byte[] FOTO7_1_SWITCH = (byte[])dt.Rows[0]["FOTO7_1_SWITCH"];
            MemoryStream mFOTO7_1_SWITCH = new MemoryStream(FOTO7_1_SWITCH);
            byte[] FOTO7_2_ROUTER = (byte[])dt.Rows[0]["FOTO7_2_ROUTER"];
            MemoryStream mFOTO7_2_ROUTER = new MemoryStream(FOTO7_2_ROUTER);
            byte[] FOTO7_3_REGLETA_ENERGIA = (byte[])dt.Rows[0]["FOTO7_3_REGLETA_ENERGIA"];
            MemoryStream mFOTO7_3_REGLETA_ENERGIA = new MemoryStream(FOTO7_3_REGLETA_ENERGIA);
            byte[] FOTO7_4_UPS = (byte[])dt.Rows[0]["FOTO7_4_UPS"];
            MemoryStream mFOTO7_4_UPS = new MemoryStream(FOTO7_4_UPS);
            byte[] FOTO7_5_COMPUTADORAS = (byte[])dt.Rows[0]["FOTO7_5_COMPUTADORAS"];
            MemoryStream mFOTO7_5_COMPUTADORAS = new MemoryStream(FOTO7_5_COMPUTADORAS);
            byte[] FOTO7_6_ACESS_POINT = (byte[])dt.Rows[0]["FOTO7_6_ACESS_POINT"];
            MemoryStream mFOTO7_6_ACESS_POINT = new MemoryStream(FOTO7_6_ACESS_POINT);
            byte[] FOTO7_7_IMPRESORA = (byte[])dt.Rows[0]["FOTO7_7_IMPRESORA"];
            MemoryStream mFOTO7_7_IMPRESORA = new MemoryStream(FOTO7_7_IMPRESORA);
            byte[] FOTO7_8_PAN_SALA_EQUIPOS = (byte[])dt.Rows[0]["FOTO7_8_PAN_SALA_EQUIPOS"];
            MemoryStream mFOTO7_8_PAN_SALA_EQUIPOS = new MemoryStream(FOTO7_8_PAN_SALA_EQUIPOS);
            byte[] FOTO7_9_JACK_RJ45 = (byte[])dt.Rows[0]["FOTO7_9_JACK_RJ45"];
            MemoryStream mFOTO7_9_JACK_RJ45 = new MemoryStream(FOTO7_9_JACK_RJ45);
            byte[] FOTO8_1_INSTALACION_POZO_TIERRA = (byte[])dt.Rows[0]["FOTO8_1_INSTALACION_POZO_TIERRA"];
            MemoryStream mFOTO8_1_INSTALACION_POZO_TIERRA = new MemoryStream(FOTO8_1_INSTALACION_POZO_TIERRA);
            byte[] FOTO8_2_CONEX_CAJA_REGISTRO = (byte[])dt.Rows[0]["FOTO8_2_CONEX_CAJA_REGISTRO"];
            MemoryStream mFOTO8_2_CONEX_CAJA_REGISTRO = new MemoryStream(FOTO8_2_CONEX_CAJA_REGISTRO);
            byte[] FOTO8_3_ESCALA_UTIL_RESULT_MEDICION1 = (byte[])dt.Rows[0]["FOTO8_3_ESCALA_UTIL_RESULT_MEDICION1"];
            MemoryStream mFOTO8_3_ESCALA_UTIL_RESULT_MEDICION1 = new MemoryStream(FOTO8_3_ESCALA_UTIL_RESULT_MEDICION1);
            byte[] FOTO8_4_ESCALA_UTIL_RESULT_MEDICION2 = (byte[])dt.Rows[0]["FOTO8_4_ESCALA_UTIL_RESULT_MEDICION2"];
            MemoryStream mFOTO8_4_ESCALA_UTIL_RESULT_MEDICION2 = new MemoryStream(FOTO8_4_ESCALA_UTIL_RESULT_MEDICION2);
            byte[] FOTO8_5_ESCALA_UTIL_RESULT_MEDICION3 = (byte[])dt.Rows[0]["FOTO8_5_ESCALA_UTIL_RESULT_MEDICION3"];
            MemoryStream mFOTO8_5_ESCALA_UTIL_RESULT_MEDICION3 = new MemoryStream(FOTO8_5_ESCALA_UTIL_RESULT_MEDICION3);
            byte[] FOTO9_1_INSTAL_POZO_TIERRA_1 = (byte[])dt.Rows[0]["FOTO9_1_INSTAL_POZO_TIERRA_1"];
            MemoryStream mFOTO9_1_INSTAL_POZO_TIERRA_1 = new MemoryStream(FOTO9_1_INSTAL_POZO_TIERRA_1);
            byte[] FOTO9_2_INSTAL_POZO_TIERRA_2 = (byte[])dt.Rows[0]["FOTO9_2_INSTAL_POZO_TIERRA_2"];
            MemoryStream mFOTO9_2_INSTAL_POZO_TIERRA_2 = new MemoryStream(FOTO9_2_INSTAL_POZO_TIERRA_2);
            byte[] FOTO9_3_ESCALA_UTIL_RESULT_MEDICION1 = (byte[])dt.Rows[0]["FOTO9_3_ESCALA_UTIL_RESULT_MEDICION1"];
            MemoryStream mFOTO9_3_ESCALA_UTIL_RESULT_MEDICION1 = new MemoryStream(FOTO9_3_ESCALA_UTIL_RESULT_MEDICION1);
            byte[] FOTO9_4_ESCALA_UTIL_RESULT_MEDICION2 = (byte[])dt.Rows[0]["FOTO9_4_ESCALA_UTIL_RESULT_MEDICION2"];
            MemoryStream mFOTO9_4_ESCALA_UTIL_RESULT_MEDICION2 = new MemoryStream(FOTO9_4_ESCALA_UTIL_RESULT_MEDICION2);
            byte[] FOTO9_5_ESCALA_UTIL_RESULT_MEDICION3 = (byte[])dt.Rows[0]["FOTO9_5_ESCALA_UTIL_RESULT_MEDICION3"];
            MemoryStream mFOTO9_5_ESCALA_UTIL_RESULT_MEDICION3 = new MemoryStream(FOTO9_5_ESCALA_UTIL_RESULT_MEDICION3);
            byte[] FOTO10_1_PANT_CONF_HOME = (byte[])dt.Rows[0]["FOTO10_1_PANT_CONF_HOME"];
            MemoryStream mFOTO10_1_PANT_CONF_HOME = new MemoryStream(FOTO10_1_PANT_CONF_HOME);
            byte[] FOTO10_2_PANT_CONF_SECURITY = (byte[])dt.Rows[0]["FOTO10_2_PANT_CONF_SECURITY"];
            MemoryStream mFOTO10_2_PANT_CONF_SECURITY = new MemoryStream(FOTO10_2_PANT_CONF_SECURITY);
            byte[] FOTO10_3_PANT_CONF_RADIO_1 = (byte[])dt.Rows[0]["FOTO10_3_PANT_CONF_RADIO_1"];
            MemoryStream mFOTO10_3_PANT_CONF_RADIO_1 = new MemoryStream(FOTO10_3_PANT_CONF_RADIO_1);
            byte[] FOTO10_4_PANT_CONF_RADIO_2 = (byte[])dt.Rows[0]["FOTO10_4_PANT_CONF_RADIO_2"];
            MemoryStream mFOTO10_4_PANT_CONF_RADIO_2 = new MemoryStream(FOTO10_4_PANT_CONF_RADIO_2);
            byte[] FOTO10_5_CONF_SISTEMA_1 = (byte[])dt.Rows[0]["FOTO10_5_CONF_SISTEMA_1"];
            MemoryStream mFOTO10_5_CONF_SISTEMA_1 = new MemoryStream(FOTO10_5_CONF_SISTEMA_1);
            byte[] FOTO10_6_CONF_SISTEMA_2 = (byte[])dt.Rows[0]["FOTO10_6_CONF_SISTEMA_2"];
            MemoryStream mFOTO10_6_CONF_SISTEMA_2 = new MemoryStream(FOTO10_6_CONF_SISTEMA_2);
            byte[] FOTO10_7_PANT_CONF_NETWORK_1 = (byte[])dt.Rows[0]["FOTO10_7_PANT_CONF_NETWORK_1"];
            MemoryStream mFOTO10_7_PANT_CONF_NETWORK_1 = new MemoryStream(FOTO10_7_PANT_CONF_NETWORK_1);
            byte[] FOTO10_8_PANT_CONF_NETWORK_2 = (byte[])dt.Rows[0]["FOTO10_8_PANT_CONF_NETWORK_2"];
            MemoryStream mFOTO10_8_PANT_CONF_NETWORK_2 = new MemoryStream(FOTO10_8_PANT_CONF_NETWORK_2);
            byte[] FOTO10_9_PANT_CONF_MONITOR_WIRELESS = (byte[])dt.Rows[0]["FOTO10_9_PANT_CONF_MONITOR_WIRELESS"];
            MemoryStream mFOTO10_9_PANT_CONF_MONITOR_WIRELESS = new MemoryStream(FOTO10_9_PANT_CONF_MONITOR_WIRELESS);
            byte[] FOTO10_10_CONF_SISTEMA_TOOLS = (byte[])dt.Rows[0]["FOTO10_10_CONF_SISTEMA_TOOLS"];
            MemoryStream mFOTO10_10_CONF_SISTEMA_TOOLS = new MemoryStream(FOTO10_10_CONF_SISTEMA_TOOLS);
            byte[] FOTO11_1_MON_CONEX_SITIO_WEB = (byte[])dt.Rows[0]["FOTO11_1_MON_CONEX_SITIO_WEB"];
            MemoryStream mFOTO11_1_MON_CONEX_SITIO_WEB = new MemoryStream(FOTO11_1_MON_CONEX_SITIO_WEB);
            byte[] FOTO11_2_MON_CONECTIVIDAD_NODO_TERMINAL = (byte[])dt.Rows[0]["FOTO11_2_MON_CONECTIVIDAD_NODO_TERMINAL"];
            MemoryStream mFOTO11_2_MON_CONECTIVIDAD_NODO_TERMINAL = new MemoryStream(FOTO11_2_MON_CONECTIVIDAD_NODO_TERMINAL);
            byte[] FOTO11_3_MON_CONECTIVIDAD_NODO_DISTRITAL = (byte[])dt.Rows[0]["FOTO11_3_MON_CONECTIVIDAD_NODO_DISTRITAL"];
            MemoryStream mFOTO11_3_MON_CONECTIVIDAD_NODO_DISTRITAL = new MemoryStream(FOTO11_3_MON_CONECTIVIDAD_NODO_DISTRITAL);
            byte[] FOTO_EPMP_1000_FORCE_180 = (byte[])dt.Rows[0]["FOTO_EPMP_1000_FORCE_180"];
            MemoryStream mFOTO_EPMP_1000_FORCE_180 = new MemoryStream(FOTO_EPMP_1000_FORCE_180);
            byte[] FOTO_1_ACCESS_POINT_SERIE = (byte[])dt.Rows[0]["FOTO_1_ACCESS_POINT_SERIE"];
            MemoryStream mFOTO_1_ACCESS_POINT_SERIE = new MemoryStream(FOTO_1_ACCESS_POINT_SERIE);
            byte[] FOTO_2_SWITCH_SERIE = (byte[])dt.Rows[0]["FOTO_2_SWITCH_SERIE"];
            MemoryStream mFOTO_2_SWITCH_SERIE = new MemoryStream(FOTO_2_SWITCH_SERIE);
            byte[] FOTO_3_ROUTER_SERIE = (byte[])dt.Rows[0]["FOTO_3_ROUTER_SERIE"];
            MemoryStream mFOTO_3_ROUTER_SERIE = new MemoryStream(FOTO_3_ROUTER_SERIE);
            byte[] FOTO_4_IMPRESORA_SERIE = (byte[])dt.Rows[0]["FOTO_4_IMPRESORA_SERIE"];
            MemoryStream mFOTO_4_IMPRESORA_SERIE = new MemoryStream(FOTO_4_IMPRESORA_SERIE);
            byte[] FOTO_5_UPS_SERIE = (byte[])dt.Rows[0]["FOTO_5_UPS_SERIE"];
            MemoryStream mFOTO_5_UPS_SERIE = new MemoryStream(FOTO_5_UPS_SERIE);
            byte[] FOTO_6_PC01_SERIE = (byte[])dt.Rows[0]["FOTO_6_PC01_SERIE"];
            MemoryStream mFOTO_6_PC01_SERIE = new MemoryStream(FOTO_6_PC01_SERIE);
            byte[] FOTO_7_PC02_SERIE = (byte[])dt.Rows[0]["FOTO_7_PC02_SERIE"];
            MemoryStream mFOTO_7_PC02_SERIE = new MemoryStream(FOTO_7_PC02_SERIE);
            byte[] FOTO_8_PC03_SERIE = (byte[])dt.Rows[0]["FOTO_8_PC03_SERIE"];
            MemoryStream mFOTO_8_PC03_SERIE = new MemoryStream(FOTO_8_PC03_SERIE);
            byte[] FOTO_9_PC04_SERIE = (byte[])dt.Rows[0]["FOTO_9_PC04_SERIE"];
            MemoryStream mFOTO_9_PC04_SERIE = new MemoryStream(FOTO_9_PC04_SERIE);
            byte[] FOTO_10_PC05_SERIE = (byte[])dt.Rows[0]["FOTO_10_PC05_SERIE"];
            MemoryStream mFOTO_10_PC05_SERIE = new MemoryStream(FOTO_10_PC05_SERIE);



            #endregion

            #endregion

            // String usuarioWindows = Environment.UserName;
            String excelGenerado = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";

            File.Copy(rutaPlantilla, excelGenerado, true);

            #region Agregando Valores por Hoja en Excel

            #region Caratula 
            ExcelToolsBL.UpdateCell(excelGenerado, "Caratula", "INSTITUCION  " + NOMBRE_IIBB, 12, "B");
            ExcelToolsBL.UpdateCell(excelGenerado, "Caratula", CODIGO_IIBB, 16, "C");
            ExcelToolsBL.UpdateCell(excelGenerado, "Caratula", TIPO_INSTITUCION, 19, "C");
            ExcelToolsBL.UpdateCell(excelGenerado, "Caratula", NOMBRE_IIBB, 22, "C");
            ExcelToolsBL.UpdateCell(excelGenerado, "Caratula", FECHA, 28, "D");
            #endregion

            #region Acta de Instalacion FITEL

            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", "PROYECTO REGIONAL DE  " + DEPARTAMENTO, 10, "C");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", NOMBRE_NODO, 14, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", UBIGEO, 14, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", DEPARTAMENTO, 15, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", PROVINCIA, 15, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", DISTRITO, 16, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", DIRECCION_NODO, 16, "J");

            if (TIPO_INSTITUCION.Equals("INSTITUCION EDUCATIVA"))
            {

                ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", "X", 19, "E");
            }
            else
            {
                if (TIPO_INSTITUCION.Equals("CENTRO DE SALUD"))
                {
                    ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", "X", 19, "H");
                }
                else
                {
                    if (TIPO_INSTITUCION.Equals("COMISARIA"))
                    {
                        ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", "X", 19, "J");
                    }
                    else
                    {
                        ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", "X", 19, "M");
                    }
                }
            }

            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", LATITUD + "º", 21, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", LONGITUD + "º", 21, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", LATITUD + "º", 21, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", DIRECCION_NODO + ", " + REFERENCIA_UBICACION_IIBB, 22, "C");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", TIPO_MASTIL, 25, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", ALTURA_MASTIL, 25, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", CODIGO_IIBB, 25, "M");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", ODU_CPE, 30, "J");

            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", SWITCH_COMUNICACIONES, 34, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", ACCESS_POINT_INDOOR, 35, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", UPS, 36, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", IMPRESORA_MULTIFUNCIONAL, 37, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", ROUTER, 38, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", EQUIPO_COMPUTO1, 39, "J");

            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", DISPONIBILIDAD_HORAS, 58, "G");

            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", POTENCIA_TRANSMISION + " DBM", 70, "F");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", FRECUENCIA + " MHz", 71, "F");
           // ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", ANCHO_BANDA_CANAL, 72, "F"); VALOR FIJO 20
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", AZIMUT + "º", 71, "L");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", FRECUENCIA, 72, "L");

            foreach (DataRow dr in dt1.Rows)
            {
                String CPE = dr["CPE"].ToString();
                String ESTACION_LOCAL = dr["ESTACION_LOCAL"].ToString();
                String RSS_CPE = dr["RSS_CPE"].ToString();
                String RSS_LOCAL = dr["RSS_LOCAL"].ToString();
                String TIEMPO_PROM = dr["TIEMPO_PROM"].ToString();
                String CAP_SUBIDA = dr["CAP_SUBIDA"].ToString();
                String CAP_BAJADA = dr["CAP_BAJADA"].ToString();
                String DISTANCIA_metros = dr["DISTANCIA_metros"].ToString();

                ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", CPE, 79, "C");
                ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", ESTACION_LOCAL, 79, "D");
                ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", RSS_CPE, 79, "E");
                ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", RSS_LOCAL, 79, "F");
                ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", TIEMPO_PROM, 79, "J");
                ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", "UL " + CAP_SUBIDA + "/" + "DL " + CAP_BAJADA, 79, "K");
                ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", DISTANCIA_metros, 79, "L");
                ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", FRECUENCIA + "MHz", 79, "M");

            }

            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", CONECTIVIDAD_GILAT, 83, "L");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", CONECTIVIDAD_NODO_TERMINAL, 84, "L");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", CONECTIVIDAD_NODO_TERMINAL, 85, "L");

            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", NOMBRES_APELLIDOS_ENCARGADO, 95, "F");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", DOC_IDENTIDAD_ENCARGADO, 96, "F");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", CELULAR_CONTACTO_ENCARGADO, 97, "F");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", EMAIL_ENCARGADO_IIBB, 98, "F");

            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", NOMBRES_APELLIDOS_REPRESENTANTE, 119, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", DOC_IDENTIDAD_REPRESENTANTE, 119, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", CARGO_REPRESENTANTE_IIBB, 119, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", EMAIL_REPRESENTANTE_IIBB, 119, "L");



            #endregion

            #region Configuracion y Pruebas

            ExcelToolsBL.UpdateCell(excelGenerado, "2 Configuración y Pruebas", NOMBRE_IIBB, 10, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "2 Configuración y Pruebas", CODIGO_IIBB, 13, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "2 Configuración y Pruebas", NOMBRE_IIBB, 14, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "2 Configuración y Pruebas", TIPO_INSTITUCION, 15, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "2 Configuración y Pruebas", LATITUD + "º", 18, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "2 Configuración y Pruebas", LONGITUD + "º", 19, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "2 Configuración y Pruebas", CODIGO_NODO, 23, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "2 Configuración y Pruebas", IP_IIBB, 27, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "2 Configuración y Pruebas", NOMBRES_APELLIDOS_REPRESENTANTE, 33, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "2 Configuración y Pruebas", CARGO_REPRESENTANTE_IIBB, 34, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "2 Configuración y Pruebas", CELULAR_CONTACTO_REPRESENTANTE, 35, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "2 Configuración y Pruebas", ODU_CPE, 38, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "2 Configuración y Pruebas", ALTURA_MASTIL, 41, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "2 Configuración y Pruebas", TIPO_MASTIL, 42, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "2 Configuración y Pruebas", ALTITUDmsnm + " M.S.N.M", 62, "D");


            #endregion

            #region Pantallas de Configuracion 

            ExcelToolsBL.AddImageDocument(false, excelGenerado, "3 Pantallas de Configuración", mPANT_CONF_ACCESS_POINt, "", 11, 3, 907, 491);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "3 Pantallas de Configuración", mPANT_CONF_ROUTER, "", 54, 3, 908, 592);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "3 Pantallas de Configuración", mPANT_CONF_SWITCH01, "", 101, 4, 768, 401);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "3 Pantallas de Configuración", mPANT_CONF_SWITCH02, "", 121, 4, 732, 339);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "3 Pantallas de Configuración", mPANT_CONF_UPS, "", 138, 4, 767, 399);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "3 Pantallas de Configuración", mPANT_CONF_ALLINONE01, "", 166, 4, 746, 447);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "3 Pantallas de Configuración", mPANT_CONF_ALLINONE01, "", 169, 4, 736, 447);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "3 Pantallas de Configuración", mPANT_CONF_IMPRESORA, "", 196, 5, 593, 399);

            #endregion

            #region Medicion SPAT   
            //ESTOS VALORES NO APLICAN SEGUN ARCHIVO SOFTWARE(1)
            //ExcelToolsBL.UpdateCell(excelGenerado, "4 Medición SPAT", MSPT_MEDIDA1_VALORMEDIO, 7, "I");
            //ExcelToolsBL.UpdateCell(excelGenerado, "4 Medición SPAT", MSPT_MEDIDA2_VALORMEDIO, 8, "I");
            //ExcelToolsBL.UpdateCell(excelGenerado, "4 Medición SPAT", MSPT_MEDIDA3_VALORMEDIO, 9, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "4 Medición SPAT", MSPTP_MEDIDA1_VALORMEDIO, 14, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "4 Medición SPAT", MSPTP_MEDIDA2_VALORMEDIO, 15, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "4 Medición SPAT", MSPTP_MEDIDA1_VALORMEDIO, 16, "I");

            #endregion

            #region Material IIBB

            ExcelToolsBL.UpdateCell(excelGenerado, "6 Material IIBB CPE", "INSTITUCION  " + NOMBRE_IIBB, 12, "F");

            foreach (DataRow dr in dt2.Rows) //DataRow dr in ds2.Tables[0].Rows)
            {
                String EQUIPO = dr["EQUIPO"].ToString();
                String MARCA = dr["MARCA"].ToString();
                String MODELO = dr["MODELO"].ToString();
                String Nro_SERIE = dr["Nro_SERIE"].ToString();

                //int ind = Convert.ToInt32(ds2.Tables[0].Rows.IndexOf(dr));
                int ind = Convert.ToInt32(dt2.Rows.IndexOf(dr));
                ExcelToolsBL.UpdateCell(excelGenerado, "6 Material IIBB CPE", Convert.ToString(ind + 1), 17 + ind, "B");
                ExcelToolsBL.UpdateCell(excelGenerado, "6 Material IIBB CPE", EQUIPO, 17 + ind, "C");
                ExcelToolsBL.UpdateCell(excelGenerado, "6 Material IIBB CPE", "1", 17 + ind, "D");
                ExcelToolsBL.UpdateCell(excelGenerado, "6 Material IIBB CPE", MARCA, 17 + ind, "E");
                ExcelToolsBL.UpdateCell(excelGenerado, "6 Material IIBB CPE", MODELO, 17 + ind, "F");
                ExcelToolsBL.UpdateCell(excelGenerado, "6 Material IIBB CPE", Nro_SERIE, 17 + ind, "G");
            }

            foreach (DataRow dr in dt3.Rows) //DataRow dr in ds3.Tables[0].Rows)
            {
                String DESCRIPCION = dr["DESCRIPCION"].ToString();
                String CODIGO = dr["CODIGO"].ToString();
                String UNIDAD = dr["UNIDAD"].ToString();
                String CANTIDAD = dr["CANTIDAD"].ToString();

                // int ind = Convert.ToInt32(ds3.Tables[0].Rows.IndexOf(dr));
                int ind = Convert.ToInt32(dt3.Rows.IndexOf(dr))
;
                ExcelToolsBL.UpdateCell(excelGenerado, "6 Material IIBB CPE", Convert.ToString(ind + 1), 32 + ind, "B");
                ExcelToolsBL.UpdateCell(excelGenerado, "6 Material IIBB CPE", DESCRIPCION, 32 + ind, "C");
                ExcelToolsBL.UpdateCell(excelGenerado, "6 Material IIBB CPE", CODIGO, 32 + ind, "D");
                ExcelToolsBL.UpdateCell(excelGenerado, "6 Material IIBB CPE", UNIDAD, 32 + ind, "F");
                ExcelToolsBL.UpdateCell(excelGenerado, "6 Material IIBB CPE", CANTIDAD, 32 + ind, "G");

            }

            #endregion

            #region Reporte Fotografico

            ExcelToolsBL.UpdateCell(excelGenerado, "7 Reporte Fotográfico IIBB CPE", CODIGO_IIBB + NOMBRE_IIBB, 7, "B");
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO1_PAN_LOCALIDAD, "", 11, 3, 904, 505);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO2_FACHADA_INSTITUCION, "", 48, 5, 984, 505);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO3_1_IMPRESORA, "", 82, 3, 404, 251);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO3_2_SWITCH, "", 82, 14, 404, 251);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO3_3_ROUTER, "", 101, 3, 405, 282);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO3_4_PC_ENCENDIDAS, "", 101, 14, 401, 284);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO3_5_PC_UPS, "", 120, 3, 404, 423);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO3_6_ACCESS_POINT, "", 120, 14, 404, 423);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO4_1_ODU_CPE, "", 141, 3, 405, 292);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO4_2_MASTIL, "", 141, 14, 405, 292);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO4_3_PAN_ANT_INSTAL_MASTIL, "", 160, 3, 401, 301);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO4_4_RECORRIDO_SFTP_CATSE, "", 160, 14, 401, 296);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO4_5_INGRESO_SFTP, "", 178, 3, 404, 282);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO4_6_RECORRIDO_SFTP_CANALETA, "", 178, 14, 401, 282);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO4_7_POE, "", 197, 3, 405, 282);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO4_8_PATCH_POE_ROUTER, "", 197, 14, 402, 280);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO5_1_TABLERO_GENERAL_SECUNDARIO, "", 219, 3, 405, 468);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO5_2_INSTALACION_BREAKER, "", 219, 14, 402, 470);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO5_3_CABLE_CONEXION_ELECTRICA, "", 237, 3, 406, 470);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO5_4_TOMAS_ENERGIA, "", 237, 14, 404, 470);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO5_5_FOTO_INTERNA_INST_BREAKER, "", 254, 14, 401, 462);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO6_1_DNI_DJREPRESENTANTE_ABONADO, "", 274, 3, 449, 408);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO6_2_DNI_DJREPRESENTANTE_ABONADO, "", 274, 13, 449, 408);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO7_1_SWITCH, "", 294, 3, 405, 280);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO7_2_ROUTER, "", 294, 14, 402, 278);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO7_3_REGLETA_ENERGIA, "", 311, 3, 405, 325);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO7_4_UPS, "", 311, 14, 403, 323);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO7_5_COMPUTADORAS, "", 328, 3, 404, 285);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO7_6_ACESS_POINT, "", 328, 14, 401, 284);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO7_7_IMPRESORA, "", 345, 3, 406, 336);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO7_8_PAN_SALA_EQUIPOS, "", 345, 14, 401, 337);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO7_9_JACK_RJ45, "", 361, 3, 405, 341);

            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO8_1_INSTALACION_POZO_TIERRA, "", 381, 3, 404, 290);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO8_2_CONEX_CAJA_REGISTRO, "", 381, 14, 403, 291);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO8_3_ESCALA_UTIL_RESULT_MEDICION1, "", 400, 3, 405, 374);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO8_4_ESCALA_UTIL_RESULT_MEDICION2, "", 400, 14, 401, 376);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO8_5_ESCALA_UTIL_RESULT_MEDICION3, "", 416, 3, 404, 375);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO9_1_INSTAL_POZO_TIERRA_1, "", 436, 3, 405, 354);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO9_2_INSTAL_POZO_TIERRA_2, "", 436, 14, 403, 354);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO9_3_ESCALA_UTIL_RESULT_MEDICION1, "", 455, 3, 404, 413);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO9_4_ESCALA_UTIL_RESULT_MEDICION2, "", 455, 14, 401, 413);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO9_5_ESCALA_UTIL_RESULT_MEDICION3, "", 471, 3, 404,412);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO10_1_PANT_CONF_HOME, "", 492, 3, 405, 286);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO10_2_PANT_CONF_SECURITY, "", 492, 14, 405, 285);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO10_3_PANT_CONF_RADIO_1, "", 509, 3, 405, 287);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO10_4_PANT_CONF_RADIO_2, "", 509, 14, 404, 286);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO10_5_CONF_SISTEMA_1, "", 526, 3, 407, 288);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO10_6_CONF_SISTEMA_2, "", 526, 14, 402, 286);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO10_7_PANT_CONF_NETWORK_1, "", 543, 3, 404, 283);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO10_8_PANT_CONF_NETWORK_2, "", 543, 14, 403, 285);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO10_9_PANT_CONF_MONITOR_WIRELESS, "", 560, 3, 407, 286);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO10_10_CONF_SISTEMA_TOOLS, "", 560, 14, 404, 285);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO11_1_MON_CONEX_SITIO_WEB, "", 580, 4, 897, 479);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO11_2_MON_CONECTIVIDAD_NODO_TERMINAL, "", 596, 3, 896, 461);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO11_3_MON_CONECTIVIDAD_NODO_DISTRITAL, "", 612, 3, 895, 459);
            #endregion

            #region Serie equipos fotos

            ExcelToolsBL.UpdateCell(excelGenerado, "6 Material IIBB CPE", "CPE: " + NOMBRE_IIBB, 13, "B");
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "8 Serie Equipos (fotos)", mFOTO_EPMP_1000_FORCE_180, "", 15, 2, 1667, 507);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "8 Serie Equipos (fotos)", mFOTO_1_ACCESS_POINT_SERIE, "", 34, 2, 709, 618);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "8 Serie Equipos (fotos)", mFOTO_2_SWITCH_SERIE, "", 34, 13, 720, 614);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "8 Serie Equipos (fotos)", mFOTO_3_ROUTER_SERIE, "", 54, 2, 707, 524);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "8 Serie Equipos (fotos)", mFOTO_4_IMPRESORA_SERIE, "", 54, 13, 720, 522);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "8 Serie Equipos (fotos)", mFOTO_5_UPS_SERIE, "", 75, 8, 609, 551);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "8 Serie Equipos (fotos)", mFOTO_6_PC01_SERIE, "", 98, 8, 609, 540);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "8 Serie Equipos (fotos)", mFOTO_7_PC02_SERIE, "", 121, 2, 707, 536);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "8 Serie Equipos (fotos)", mFOTO_8_PC03_SERIE, "", 121, 13, 722, 538);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "8 Serie Equipos (fotos)", mFOTO_9_PC04_SERIE, "", 141, 2, 707, 526);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "8 Serie Equipos (fotos)", mFOTO_10_PC05_SERIE, "", 141, 13, 722, 524);


            #endregion

            #endregion

            #region Ruta del Zip
            String rutaNodo = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo;
            if (Directory.Exists(rutaNodo))
            {
                String rutaParcial = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\CS01\\";
                if (!Directory.Exists(rutaParcial)) Directory.CreateDirectory(rutaParcial);
                String rutaAlterna = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\CS01\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";
                File.Copy(excelGenerado, rutaAlterna, true);
            }
            #endregion
        }
        public void ActaInstalacionAceptacionProtocoloIIBB_B(String IdNodo, String IdTarea, String valorCadena1, String rutaPlantilla)
      {
            DBBaseDatos baseDatosDA = new DBBaseDatos();

            baseDatosDA.Configurar();
            baseDatosDA.Conectar();
            DataTable dt = new DataTable();
            DataTable dt1 = new DataTable();
            DataTable dt2 = new DataTable();
            DataTable dt3 = new DataTable();
            try
            {
                baseDatosDA.CrearComando("USP_R_ACTA_INSTALACION_ACEPTACION_PROTOCOLO_IIBB_B", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA",IdTarea,true);
                dt = baseDatosDA.EjecutarConsultaDataTable();

                baseDatosDA.CrearComando("USP_R_MEDICION_ENLACE_IIBB_B", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                dt1 = baseDatosDA.EjecutarConsultaDataTable();

                baseDatosDA.CrearComando("USP_R_EQUIPAMIENTOS_IIBB_B", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                dt2 = baseDatosDA.EjecutarConsultaDataTable();

                baseDatosDA.CrearComando("USP_R_MATERIALES_IIBB_B", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                dt3 = baseDatosDA.EjecutarConsultaDataTable();

              

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
            #region Valores

            #region Valores String
            String FRECUENCIA = dt.Rows[0]["FRECUENCIA"].ToString();
            String CODIGO_IIBB = dt.Rows[0]["CODIGO_IIBB"].ToString();
            String TIPO_INSTITUCION = dt.Rows[0]["TIPO_INSTITUCION"].ToString();
            String NOMBRE_IIBB = dt.Rows[0]["NOMBRE_IIBB"].ToString();

            String fechaSQL = dt.Rows[0]["FECHA"].ToString();
            String FECHA = "";
            if (fechaSQL != "")
            {
                DateTime dtFecha = DateTime.Parse(fechaSQL);
                FECHA = dtFecha.ToString("dd/MM/yyyy");
            }
            else { FECHA = ""; }

            String DEPARTAMENTO = dt.Rows[0]["DEPARTAMENTO"].ToString();
            String PROVINCIA = dt.Rows[0]["PROVINCIA"].ToString();
            String DISTRITO = dt.Rows[0]["DISTRITO"].ToString();
            String NOMBRE_NODO = dt.Rows[0]["NOMBRE_NODO"].ToString();
            String UBIGEO = dt.Rows[0]["UBIGEO"].ToString();
            String LATITUD = dt.Rows[0]["LATITUD"].ToString();
            String LONGITUD = dt.Rows[0]["LONGITUD"].ToString();
            String ALTITUDmsnm = dt.Rows[0]["ALTITUDmsnm"].ToString();
            String AZIMUT = dt.Rows[0]["AZIMUT"].ToString();


            String DIRECCION_NODO = dt.Rows[0]["DIRECCION_NODO"].ToString();

            String ODU_CPE = dt.Rows[0]["ODU_CPE"].ToString();
            String ACCESS_POINT_INDOOR = dt.Rows[0]["ACCESS_POINT_INDOOR"].ToString();
            String SWITCH_COMUNICACIONES = dt.Rows[0]["SWITCH_COMUNICACIONES"].ToString();
            String ROUTER = dt.Rows[0]["ROUTER"].ToString();
            String EQUIPO_COMPUTO1 = dt.Rows[0]["EQUIPO_COMPUTO1"].ToString();
            String EQUIPO_COMPUTO2 = dt.Rows[0]["EQUIPO_COMPUTO2"].ToString();
            String EQUIPO_COMPUTO3 = dt.Rows[0]["EQUIPO_COMPUTO3"].ToString();
            String EQUIPO_COMPUTO4 = dt.Rows[0]["EQUIPO_COMPUTO4"].ToString();
            String EQUIPO_COMPUTO5 = dt.Rows[0]["EQUIPO_COMPUTO5"].ToString();
            String IMPRESORA_MULTIFUNCIONAL = dt.Rows[0]["IMPRESORA_MULTIFUNCIONAL"].ToString();
            String UPS = dt.Rows[0]["UPS"].ToString();
            String REFERENCIA_UBICACION_IIBB = dt.Rows[0]["REFERENCIA_UBICACION_IIBB"].ToString();
            String TIPO_MASTIL = dt.Rows[0]["TIPO_MASTIL"].ToString();
            String ALTURA_MASTIL = dt.Rows[0]["ALTURA_MASTIL"].ToString();
            String DISPONIBILIDAD_HORAS = dt.Rows[0]["DISPONIBILIDAD_HORAS"].ToString();
            String VALOR_MEDIO_MEDIDA1 = dt.Rows[0]["VALOR_MEDIO_MEDIDA1"].ToString();
            String VALOR_MEDIO_MEDIDA2 = dt.Rows[0]["VALOR_MEDIO_MEDIDA2"].ToString();
            String VALOR_MEDIO_MEDIDA3 = dt.Rows[0]["VALOR_MEDIO_MEDIDA3"].ToString();
            String POTENCIA_TRANSMISION = dt.Rows[0]["POTENCIA_TRANSMISION"].ToString();
            String ANCHO_BANDA_CANAL = dt.Rows[0]["ANCHO_BANDA_CANAL"].ToString();    
            String ELEVACION = dt.Rows[0]["ELEVACION"].ToString();
            String CONECTIVIDAD_GILAT = dt.Rows[0]["CONECTIVIDAD_GILAT"].ToString();
            String CONECTIVIDAD_NODO_TERMINAL = dt.Rows[0]["CONECTIVIDAD_NODO_TERMINAL"].ToString();
            String CONECTIVIDAD_NODO_DISTRITAL = dt.Rows[0]["CONECTIVIDAD_NODO_DISTRITAL"].ToString();
            String CONECTIVIDAD_NOC = dt.Rows[0]["CONECTIVIDAD_NOC"].ToString();
            String NOMBRES_APELLIDOS_ENCARGADO = dt.Rows[0]["NOMBRES_APELLIDOS_ENCARGADO"].ToString();
            String DOC_IDENTIDAD_ENCARGADO = dt.Rows[0]["DOC_IDENTIDAD_ENCARGADO"].ToString();
            String CELULAR_CONTACTO_ENCARGADO = dt.Rows[0]["CELULAR_CONTACTO_ENCARGADO"].ToString();
            String EMAIL_ENCARGADO_IIBB = dt.Rows[0]["EMAIL_ENCARGADO_IIBB"].ToString();
            String NOMBRES_APELLIDOS_REPRESENTANTE = dt.Rows[0]["NOMBRES_APELLIDOS_REPRESENTANTE"].ToString();
            String DOC_IDENTIDAD_REPRESENTANTE = dt.Rows[0]["DOC_IDENTIDAD_REPRESENTANTE"].ToString();
            String CELULAR_CONTACTO_REPRESENTANTE = dt.Rows[0]["CELULAR_CONTACTO_REPRESENTANTE"].ToString();
            String CARGO_REPRESENTANTE_IIBB = dt.Rows[0]["CARGO_REPRESENTANTE_IIBB"].ToString();
            String EMAIL_REPRESENTANTE_IIBB = dt.Rows[0]["EMAIL_REPRESENTANTE_IIBB"].ToString();
            String NOMBRES_APELLIDOS_REPR_OPERADOR = dt.Rows[0]["NOMBRES_APELLIDOS_REPR_OPERADOR"].ToString();
            String DOC_IDENTIDAD_REPR_OPERADOR = dt.Rows[0]["DOC_IDENTIDAD_REPR_OPERADOR"].ToString();
            String CARGO_REPRESENTANTE_OPERADOR = dt.Rows[0]["CARGO_REPRESENTANTE_OPERADOR"].ToString();
            String EMAIL_REPR_OPERADOR = dt.Rows[0]["EMAIL_REPR_OPERADOR"].ToString();

            String MSPT_MEDIDA1_VALORMEDIO = dt.Rows[0]["MSPT_MEDIDA1_VALORMEDIO"].ToString();
            String MSPT_MEDIDA2_VALORMEDIO = dt.Rows[0]["MSPT_MEDIDA2_VALORMEDIO"].ToString();
            String MSPT_MEDIDA3_VALORMEDIO = dt.Rows[0]["MSPT_MEDIDA3_VALORMEDIO"].ToString();
            String MSPTP_MEDIDA1_VALORMEDIO = dt.Rows[0]["MSPTP_MEDIDA1_VALORMEDIO"].ToString();
            String MSPTP_MEDIDA2_VALORMEDIO = dt.Rows[0]["MSPTP_MEDIDA2_VALORMEDIO"].ToString();
            String MSPTP_MEDIDA3_VALORMEDIO = dt.Rows[0]["MSPTP_MEDIDA3_VALORMEDIO"].ToString();

            String CODIGO_NODO = dt.Rows[0]["CODIGO_NODO"].ToString();
            String IP_IIBB = dt.Rows[0]["IP_IIBB"].ToString();

            #endregion

            #region Valores byte

            byte[] PANT_CONF_ACCESS_POINt = (byte[])dt.Rows[0]["PANT_CONF_ACCESS_POINt"];
            MemoryStream mPANT_CONF_ACCESS_POINt = new MemoryStream(PANT_CONF_ACCESS_POINt);
            byte[] PANT_CONF_ROUTER = (byte[])dt.Rows[0]["PANT_CONF_ROUTER"];
            MemoryStream mPANT_CONF_ROUTER = new MemoryStream(PANT_CONF_ROUTER);
            byte[] PANT_CONF_SWITCH01 = (byte[])dt.Rows[0]["PANT_CONF_SWITCH01"];
            MemoryStream mPANT_CONF_SWITCH01 = new MemoryStream(PANT_CONF_SWITCH01);
            byte[] PANT_CONF_SWITCH02 = (byte[])dt.Rows[0]["PANT_CONF_SWITCH02"];
            MemoryStream mPANT_CONF_SWITCH02 = new MemoryStream(PANT_CONF_SWITCH02);
            byte[] PANT_CONF_UPS = (byte[])dt.Rows[0]["PANT_CONF_UPS"];
            MemoryStream mPANT_CONF_UPS = new MemoryStream(PANT_CONF_UPS);
            byte[] PANT_CONF_ALLINONE01 = (byte[])dt.Rows[0]["PANT_CONF_ALLINONE01"];
            MemoryStream mPANT_CONF_ALLINONE01 = new MemoryStream(PANT_CONF_ALLINONE01);
            byte[] PANT_CONF_ALLINONE02 = (byte[])dt.Rows[0]["PANT_CONF_ALLINONE02"];
            MemoryStream mPANT_CONF_ALLINONE02 = new MemoryStream(PANT_CONF_ALLINONE02);
            byte[] PANT_CONF_IMPRESORA = (byte[])dt.Rows[0]["PANT_CONF_IMPRESORA"];
            MemoryStream mPANT_CONF_IMPRESORA = new MemoryStream(PANT_CONF_IMPRESORA);
            byte[] FOTO1_PAN_LOCALIDAD = (byte[])dt.Rows[0]["FOTO1_PAN_LOCALIDAD"];
            MemoryStream mFOTO1_PAN_LOCALIDAD = new MemoryStream(FOTO1_PAN_LOCALIDAD);
            byte[] FOTO2_FACHADA_INSTITUCION = (byte[])dt.Rows[0]["FOTO2_FACHADA_INSTITUCION"];
            MemoryStream mFOTO2_FACHADA_INSTITUCION = new MemoryStream(FOTO2_FACHADA_INSTITUCION);
            byte[] FOTO3_1_IMPRESORA = (byte[])dt.Rows[0]["FOTO3_1_IMPRESORA"];
            MemoryStream mFOTO3_1_IMPRESORA = new MemoryStream(FOTO3_1_IMPRESORA);
            byte[] FOTO3_2_SWITCH = (byte[])dt.Rows[0]["FOTO3_2_SWITCH"];
            MemoryStream mFOTO3_2_SWITCH = new MemoryStream(FOTO3_2_SWITCH);
            byte[] FOTO3_3_ROUTER = (byte[])dt.Rows[0]["FOTO3_3_ROUTER"];
            MemoryStream mFOTO3_3_ROUTER = new MemoryStream(FOTO3_3_ROUTER);
            byte[] FOTO3_4_PC_ENCENDIDAS = (byte[])dt.Rows[0]["FOTO3_4_PC_ENCENDIDAS"];
            MemoryStream mFOTO3_4_PC_ENCENDIDAS = new MemoryStream(FOTO3_4_PC_ENCENDIDAS);
            byte[] FOTO3_5_PC_UPS = (byte[])dt.Rows[0]["FOTO3_5_PC_UPS"];
            MemoryStream mFOTO3_5_PC_UPS = new MemoryStream(FOTO3_5_PC_UPS);
            byte[] FOTO3_6_ACCESS_POINT = (byte[])dt.Rows[0]["FOTO3_6_ACCESS_POINT"];
            MemoryStream mFOTO3_6_ACCESS_POINT = new MemoryStream(FOTO3_6_ACCESS_POINT);
            byte[] FOTO4_1_ODU_CPE = (byte[])dt.Rows[0]["FOTO4_1_ODU_CPE"];
            MemoryStream mFOTO4_1_ODU_CPE = new MemoryStream(FOTO4_1_ODU_CPE);
            byte[] FOTO4_2_MASTIL = (byte[])dt.Rows[0]["FOTO4_2_MASTIL"];
            MemoryStream mFOTO4_2_MASTIL = new MemoryStream(FOTO4_2_MASTIL);
            byte[] FOTO4_3_PAN_ANT_INSTAL_MASTIL = (byte[])dt.Rows[0]["FOTO4_3_PAN_ANT_INSTAL_MASTIL"];
            MemoryStream mFOTO4_3_PAN_ANT_INSTAL_MASTIL = new MemoryStream(FOTO4_3_PAN_ANT_INSTAL_MASTIL);
            byte[] FOTO4_4_RECORRIDO_SFTP_CATSE = (byte[])dt.Rows[0]["FOTO4_4_RECORRIDO_SFTP_CATSE"];
            MemoryStream mFOTO4_4_RECORRIDO_SFTP_CATSE = new MemoryStream(FOTO4_4_RECORRIDO_SFTP_CATSE);
            byte[] FOTO4_5_INGRESO_SFTP = (byte[])dt.Rows[0]["FOTO4_5_INGRESO_SFTP"];
            MemoryStream mFOTO4_5_INGRESO_SFTP = new MemoryStream(FOTO4_5_INGRESO_SFTP);
            byte[] FOTO4_6_RECORRIDO_SFTP_CANALETA = (byte[])dt.Rows[0]["FOTO4_6_RECORRIDO_SFTP_CANALETA"];
            MemoryStream mFOTO4_6_RECORRIDO_SFTP_CANALETA = new MemoryStream(FOTO4_6_RECORRIDO_SFTP_CANALETA);
            byte[] FOTO4_7_POE = (byte[])dt.Rows[0]["FOTO4_7_POE"];
            MemoryStream mFOTO4_7_POE = new MemoryStream(FOTO4_7_POE);
            byte[] FOTO4_8_PATCH_POE_ROUTER = (byte[])dt.Rows[0]["FOTO4_8_PATCH_POE_ROUTER"];
            MemoryStream mFOTO4_8_PATCH_POE_ROUTER = new MemoryStream(FOTO4_8_PATCH_POE_ROUTER);
            byte[] FOTO5_1_TABLERO_GENERAL_SECUNDARIO = (byte[])dt.Rows[0]["FOTO5_1_TABLERO_GENERAL_SECUNDARIO"];
            MemoryStream mFOTO5_1_TABLERO_GENERAL_SECUNDARIO = new MemoryStream(FOTO5_1_TABLERO_GENERAL_SECUNDARIO);
            byte[] FOTO5_2_INSTALACION_BREAKER = (byte[])dt.Rows[0]["FOTO5_2_INSTALACION_BREAKER"];
            MemoryStream mFOTO5_2_INSTALACION_BREAKER = new MemoryStream(FOTO5_2_INSTALACION_BREAKER);
            byte[] FOTO5_3_CABLE_CONEXION_ELECTRICA = (byte[])dt.Rows[0]["FOTO5_3_CABLE_CONEXION_ELECTRICA"];
            MemoryStream mFOTO5_3_CABLE_CONEXION_ELECTRICA = new MemoryStream(FOTO5_3_CABLE_CONEXION_ELECTRICA);
            byte[] FOTO5_4_TOMAS_ENERGIA = (byte[])dt.Rows[0]["FOTO5_4_TOMAS_ENERGIA"];
            MemoryStream mFOTO5_4_TOMAS_ENERGIA = new MemoryStream(FOTO5_4_TOMAS_ENERGIA);
            byte[] FOTO5_5_FOTO_INTERNA_INST_BREAKER = (byte[])dt.Rows[0]["FOTO5_5_FOTO_INTERNA_INST_BREAKER"];
            MemoryStream mFOTO5_5_FOTO_INTERNA_INST_BREAKER = new MemoryStream(FOTO5_5_FOTO_INTERNA_INST_BREAKER);
            byte[] FOTO6_1_DNI_DJREPRESENTANTE_ABONADO = (byte[])dt.Rows[0]["FOTO6_1_DNI_DJREPRESENTANTE_ABONADO"];
            MemoryStream mFOTO6_1_DNI_DJREPRESENTANTE_ABONADO = new MemoryStream(FOTO6_1_DNI_DJREPRESENTANTE_ABONADO);
            byte[] FOTO6_2_DNI_DJREPRESENTANTE_ABONADO = (byte[])dt.Rows[0]["FOTO6_2_DNI_DJREPRESENTANTE_ABONADO"];
            MemoryStream mFOTO6_2_DNI_DJREPRESENTANTE_ABONADO = new MemoryStream(FOTO6_2_DNI_DJREPRESENTANTE_ABONADO);
            byte[] FOTO7_1_SWITCH = (byte[])dt.Rows[0]["FOTO7_1_SWITCH"];
            MemoryStream mFOTO7_1_SWITCH = new MemoryStream(FOTO7_1_SWITCH);
            byte[] FOTO7_2_ROUTER = (byte[])dt.Rows[0]["FOTO7_2_ROUTER"];
            MemoryStream mFOTO7_2_ROUTER = new MemoryStream(FOTO7_2_ROUTER);
            byte[] FOTO7_3_REGLETA_ENERGIA = (byte[])dt.Rows[0]["FOTO7_3_REGLETA_ENERGIA"];
            MemoryStream mFOTO7_3_REGLETA_ENERGIA = new MemoryStream(FOTO7_3_REGLETA_ENERGIA);
            byte[] FOTO7_4_UPS = (byte[])dt.Rows[0]["FOTO7_4_UPS"];
            MemoryStream mFOTO7_4_UPS = new MemoryStream(FOTO7_4_UPS);
            byte[] FOTO7_5_COMPUTADORAS = (byte[])dt.Rows[0]["FOTO7_5_COMPUTADORAS"];
            MemoryStream mFOTO7_5_COMPUTADORAS = new MemoryStream(FOTO7_5_COMPUTADORAS);
            byte[] FOTO7_6_ACESS_POINT = (byte[])dt.Rows[0]["FOTO7_6_ACESS_POINT"];
            MemoryStream mFOTO7_6_ACESS_POINT = new MemoryStream(FOTO7_6_ACESS_POINT);
            byte[] FOTO7_7_IMPRESORA = (byte[])dt.Rows[0]["FOTO7_7_IMPRESORA"];
            MemoryStream mFOTO7_7_IMPRESORA = new MemoryStream(FOTO7_7_IMPRESORA);
            byte[] FOTO7_8_PAN_SALA_EQUIPOS = (byte[])dt.Rows[0]["FOTO7_8_PAN_SALA_EQUIPOS"];
            MemoryStream mFOTO7_8_PAN_SALA_EQUIPOS = new MemoryStream(FOTO7_8_PAN_SALA_EQUIPOS);
            byte[] FOTO7_9_JACK_RJ45 = (byte[])dt.Rows[0]["FOTO7_9_JACK_RJ45"];
            MemoryStream mFOTO7_9_JACK_RJ45 = new MemoryStream(FOTO7_9_JACK_RJ45);
            byte[] FOTO8_1_INSTALACION_POZO_TIERRA = (byte[])dt.Rows[0]["FOTO8_1_INSTALACION_POZO_TIERRA"];
            MemoryStream mFOTO8_1_INSTALACION_POZO_TIERRA = new MemoryStream(FOTO8_1_INSTALACION_POZO_TIERRA);
            byte[] FOTO8_2_CONEX_CAJA_REGISTRO = (byte[])dt.Rows[0]["FOTO8_2_CONEX_CAJA_REGISTRO"];
            MemoryStream mFOTO8_2_CONEX_CAJA_REGISTRO = new MemoryStream(FOTO8_2_CONEX_CAJA_REGISTRO);
            byte[] FOTO8_3_ESCALA_UTIL_RESULT_MEDICION1 = (byte[])dt.Rows[0]["FOTO8_3_ESCALA_UTIL_RESULT_MEDICION1"];
            MemoryStream mFOTO8_3_ESCALA_UTIL_RESULT_MEDICION1 = new MemoryStream(FOTO8_3_ESCALA_UTIL_RESULT_MEDICION1);
            byte[] FOTO8_4_ESCALA_UTIL_RESULT_MEDICION2 = (byte[])dt.Rows[0]["FOTO8_4_ESCALA_UTIL_RESULT_MEDICION2"];
            MemoryStream mFOTO8_4_ESCALA_UTIL_RESULT_MEDICION2 = new MemoryStream(FOTO8_4_ESCALA_UTIL_RESULT_MEDICION2);
            byte[] FOTO8_5_ESCALA_UTIL_RESULT_MEDICION3 = (byte[])dt.Rows[0]["FOTO8_5_ESCALA_UTIL_RESULT_MEDICION3"];
            MemoryStream mFOTO8_5_ESCALA_UTIL_RESULT_MEDICION3 = new MemoryStream(FOTO8_5_ESCALA_UTIL_RESULT_MEDICION3);
            byte[] FOTO9_1_INSTAL_POZO_TIERRA_1 = (byte[])dt.Rows[0]["FOTO9_1_INSTAL_POZO_TIERRA_1"];
            MemoryStream mFOTO9_1_INSTAL_POZO_TIERRA_1 = new MemoryStream(FOTO9_1_INSTAL_POZO_TIERRA_1);
            byte[] FOTO9_2_INSTAL_POZO_TIERRA_2 = (byte[])dt.Rows[0]["FOTO9_2_INSTAL_POZO_TIERRA_2"];
            MemoryStream mFOTO9_2_INSTAL_POZO_TIERRA_2 = new MemoryStream(FOTO9_2_INSTAL_POZO_TIERRA_2);
            byte[] FOTO9_3_ESCALA_UTIL_RESULT_MEDICION1 = (byte[])dt.Rows[0]["FOTO9_3_ESCALA_UTIL_RESULT_MEDICION1"];
            MemoryStream mFOTO9_3_ESCALA_UTIL_RESULT_MEDICION1 = new MemoryStream(FOTO9_3_ESCALA_UTIL_RESULT_MEDICION1);
            byte[] FOTO9_4_ESCALA_UTIL_RESULT_MEDICION2 = (byte[])dt.Rows[0]["FOTO9_4_ESCALA_UTIL_RESULT_MEDICION2"];
            MemoryStream mFOTO9_4_ESCALA_UTIL_RESULT_MEDICION2 = new MemoryStream(FOTO9_4_ESCALA_UTIL_RESULT_MEDICION2);
            byte[] FOTO9_5_ESCALA_UTIL_RESULT_MEDICION3 = (byte[])dt.Rows[0]["FOTO9_5_ESCALA_UTIL_RESULT_MEDICION3"];
            MemoryStream mFOTO9_5_ESCALA_UTIL_RESULT_MEDICION3 = new MemoryStream(FOTO9_5_ESCALA_UTIL_RESULT_MEDICION3);
            byte[] FOTO10_1_PANT_CONF_HOME = (byte[])dt.Rows[0]["FOTO10_1_PANT_CONF_HOME"];
            MemoryStream mFOTO10_1_PANT_CONF_HOME = new MemoryStream(FOTO10_1_PANT_CONF_HOME);
            byte[] FOTO10_2_PANT_CONF_SECURITY = (byte[])dt.Rows[0]["FOTO10_2_PANT_CONF_SECURITY"];
            MemoryStream mFOTO10_2_PANT_CONF_SECURITY = new MemoryStream(FOTO10_2_PANT_CONF_SECURITY);
            byte[] FOTO10_3_PANT_CONF_RADIO_1 = (byte[])dt.Rows[0]["FOTO10_3_PANT_CONF_RADIO_1"];
            MemoryStream mFOTO10_3_PANT_CONF_RADIO_1 = new MemoryStream(FOTO10_3_PANT_CONF_RADIO_1);
            byte[] FOTO10_4_PANT_CONF_RADIO_2 = (byte[])dt.Rows[0]["FOTO10_4_PANT_CONF_RADIO_2"];
            MemoryStream mFOTO10_4_PANT_CONF_RADIO_2 = new MemoryStream(FOTO10_4_PANT_CONF_RADIO_2);
            byte[] FOTO10_5_CONF_SISTEMA_1 = (byte[])dt.Rows[0]["FOTO10_5_CONF_SISTEMA_1"];
            MemoryStream mFOTO10_5_CONF_SISTEMA_1 = new MemoryStream(FOTO10_5_CONF_SISTEMA_1);
            byte[] FOTO10_6_CONF_SISTEMA_2 = (byte[])dt.Rows[0]["FOTO10_6_CONF_SISTEMA_2"];
            MemoryStream mFOTO10_6_CONF_SISTEMA_2 = new MemoryStream(FOTO10_6_CONF_SISTEMA_2);
            byte[] FOTO10_7_PANT_CONF_NETWORK_1 = (byte[])dt.Rows[0]["FOTO10_7_PANT_CONF_NETWORK_1"];
            MemoryStream mFOTO10_7_PANT_CONF_NETWORK_1 = new MemoryStream(FOTO10_7_PANT_CONF_NETWORK_1);
            byte[] FOTO10_8_PANT_CONF_NETWORK_2 = (byte[])dt.Rows[0]["FOTO10_8_PANT_CONF_NETWORK_2"];
            MemoryStream mFOTO10_8_PANT_CONF_NETWORK_2 = new MemoryStream(FOTO10_8_PANT_CONF_NETWORK_2);
            byte[] FOTO10_9_PANT_CONF_MONITOR_WIRELESS = (byte[])dt.Rows[0]["FOTO10_9_PANT_CONF_MONITOR_WIRELESS"];
            MemoryStream mFOTO10_9_PANT_CONF_MONITOR_WIRELESS = new MemoryStream(FOTO10_9_PANT_CONF_MONITOR_WIRELESS);
            byte[] FOTO10_10_CONF_SISTEMA_TOOLS = (byte[])dt.Rows[0]["FOTO10_10_CONF_SISTEMA_TOOLS"];
            MemoryStream mFOTO10_10_CONF_SISTEMA_TOOLS = new MemoryStream(FOTO10_10_CONF_SISTEMA_TOOLS);
            byte[] FOTO11_1_MON_CONEX_SITIO_WEB = (byte[])dt.Rows[0]["FOTO11_1_MON_CONEX_SITIO_WEB"];
            MemoryStream mFOTO11_1_MON_CONEX_SITIO_WEB = new MemoryStream(FOTO11_1_MON_CONEX_SITIO_WEB);
            byte[] FOTO11_2_MON_CONECTIVIDAD_NODO_TERMINAL = (byte[])dt.Rows[0]["FOTO11_2_MON_CONECTIVIDAD_NODO_TERMINAL"];
            MemoryStream mFOTO11_2_MON_CONECTIVIDAD_NODO_TERMINAL = new MemoryStream(FOTO11_2_MON_CONECTIVIDAD_NODO_TERMINAL);
            byte[] FOTO11_3_MON_CONECTIVIDAD_NODO_DISTRITAL = (byte[])dt.Rows[0]["FOTO11_3_MON_CONECTIVIDAD_NODO_DISTRITAL"];
            MemoryStream mFOTO11_3_MON_CONECTIVIDAD_NODO_DISTRITAL = new MemoryStream(FOTO11_3_MON_CONECTIVIDAD_NODO_DISTRITAL);
            byte[] FOTO_EPMP_1000_FORCE_180 = (byte[])dt.Rows[0]["FOTO_EPMP_1000_FORCE_180"];
            MemoryStream mFOTO_EPMP_1000_FORCE_180 = new MemoryStream(FOTO_EPMP_1000_FORCE_180);
            byte[] FOTO_1_ACCESS_POINT_SERIE = (byte[])dt.Rows[0]["FOTO_1_ACCESS_POINT_SERIE"];
            MemoryStream mFOTO_1_ACCESS_POINT_SERIE = new MemoryStream(FOTO_1_ACCESS_POINT_SERIE);
            byte[] FOTO_2_SWITCH_SERIE = (byte[])dt.Rows[0]["FOTO_2_SWITCH_SERIE"];
            MemoryStream mFOTO_2_SWITCH_SERIE = new MemoryStream(FOTO_2_SWITCH_SERIE);
            byte[] FOTO_3_ROUTER_SERIE = (byte[])dt.Rows[0]["FOTO_3_ROUTER_SERIE"];
            MemoryStream mFOTO_3_ROUTER_SERIE = new MemoryStream(FOTO_3_ROUTER_SERIE);
            byte[] FOTO_4_IMPRESORA_SERIE = (byte[])dt.Rows[0]["FOTO_4_IMPRESORA_SERIE"];
            MemoryStream mFOTO_4_IMPRESORA_SERIE = new MemoryStream(FOTO_4_IMPRESORA_SERIE);
            byte[] FOTO_5_UPS_SERIE = (byte[])dt.Rows[0]["FOTO_5_UPS_SERIE"];
            MemoryStream mFOTO_5_UPS_SERIE = new MemoryStream(FOTO_5_UPS_SERIE);
            byte[] FOTO_6_PC01_SERIE = (byte[])dt.Rows[0]["FOTO_6_PC01_SERIE"];
            MemoryStream mFOTO_6_PC01_SERIE = new MemoryStream(FOTO_6_PC01_SERIE);
            byte[] FOTO_7_PC02_SERIE = (byte[])dt.Rows[0]["FOTO_7_PC02_SERIE"];
            MemoryStream mFOTO_7_PC02_SERIE = new MemoryStream(FOTO_7_PC02_SERIE);
            byte[] FOTO_8_PC03_SERIE = (byte[])dt.Rows[0]["FOTO_8_PC03_SERIE"];
            MemoryStream mFOTO_8_PC03_SERIE = new MemoryStream(FOTO_8_PC03_SERIE);
            byte[] FOTO_9_PC04_SERIE = (byte[])dt.Rows[0]["FOTO_9_PC04_SERIE"];
            MemoryStream mFOTO_9_PC04_SERIE = new MemoryStream(FOTO_9_PC04_SERIE);
            byte[] FOTO_10_PC05_SERIE = (byte[])dt.Rows[0]["FOTO_10_PC05_SERIE"];
            MemoryStream mFOTO_10_PC05_SERIE = new MemoryStream(FOTO_10_PC05_SERIE);

            #endregion

            #endregion

            String excelGenerado = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";

            File.Copy(rutaPlantilla, excelGenerado, true);

            #region Agregando Valores por Hoja en Excel

            #region Caratula 
            ExcelToolsBL.UpdateCell(excelGenerado, "Caratula", "INSTITUCION  " + NOMBRE_IIBB, 12, "B");
            ExcelToolsBL.UpdateCell(excelGenerado, "Caratula", CODIGO_IIBB, 16, "C");
            ExcelToolsBL.UpdateCell(excelGenerado, "Caratula", TIPO_INSTITUCION, 19, "C");
            ExcelToolsBL.UpdateCell(excelGenerado, "Caratula", NOMBRE_IIBB, 22, "C");
            ExcelToolsBL.UpdateCell(excelGenerado, "Caratula", FECHA, 28, "D");
            #endregion

            #region Acta de Instalacion FITEL

            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", "PROYECTO REGIONAL DE  " + DEPARTAMENTO, 10, "C");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", NOMBRE_NODO, 14, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", UBIGEO, 14, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", DEPARTAMENTO, 15, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", PROVINCIA, 15, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", DISTRITO, 16, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", DIRECCION_NODO, 16, "J");

            if (TIPO_INSTITUCION.Equals("INSTITUCION EDUCATIVA"))
            {

                ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", "X", 19, "E");
            }
            else
            {
                if (TIPO_INSTITUCION.Equals("CENTRO DE SALUD"))
                {
                    ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", "X", 19, "H");
                }
                else
                {
                    if (TIPO_INSTITUCION.Equals("COMISARIA"))
                    {
                        ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", "X", 19, "J");
                    }
                    else
                    {
                        ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", "X", 19, "M");
                    }
                }
            }

            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", LATITUD + "º", 21, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", LONGITUD + "º", 21, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", LATITUD + "º", 21, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", DIRECCION_NODO + ", " + REFERENCIA_UBICACION_IIBB, 22, "C");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", TIPO_MASTIL, 25, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", ALTURA_MASTIL, 25, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", CODIGO_IIBB, 25, "M");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", ODU_CPE, 30, "J");

            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", SWITCH_COMUNICACIONES, 34, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", ACCESS_POINT_INDOOR, 35, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", UPS, 36, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", IMPRESORA_MULTIFUNCIONAL, 37, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", ROUTER, 38, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", EQUIPO_COMPUTO1, 39, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", EQUIPO_COMPUTO2, 40, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", EQUIPO_COMPUTO3, 41, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", EQUIPO_COMPUTO4, 42, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", EQUIPO_COMPUTO5, 43, "J");

            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", DISPONIBILIDAD_HORAS, 58, "G");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", VALOR_MEDIO_MEDIDA1, 62, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", VALOR_MEDIO_MEDIDA2, 63, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", VALOR_MEDIO_MEDIDA3, 64, "J");

            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", POTENCIA_TRANSMISION + " DBM", 70, "F");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", FRECUENCIA + " MHz", 71, "F");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", ANCHO_BANDA_CANAL, 72, "F");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", AZIMUT + "º", 71, "L");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", FRECUENCIA, 72, "L");

            foreach (DataRow dr in dt1.Rows)
            {
                String CPE = dr["CPE"].ToString();
                String ESTACION_LOCAL = dr["ESTACION_LOCAL"].ToString();
                String RSS_CPE = dr["RSS_CPE"].ToString();
                String RSS_LOCAL = dr["RSS_LOCAL"].ToString();
                String TIEMPO_PROM = dr["TIEMPO_PROM"].ToString();
                String CAP_SUBIDA = dr["CAP_SUBIDA"].ToString();
                String CAP_BAJADA = dr["CAP_BAJADA"].ToString();
                String DISTANCIA_metros = dr["DISTANCIA_metros"].ToString();

                ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", CPE, 79, "C");
                ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", ESTACION_LOCAL, 79, "D");
                ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", RSS_CPE, 79, "E");
                ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", RSS_LOCAL, 79, "F");
                ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", TIEMPO_PROM, 79, "J");
                ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", "UL " + CAP_SUBIDA + "/" + "DL " + CAP_BAJADA, 79, "K");
                ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", DISTANCIA_metros, 79, "L");
                ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", FRECUENCIA + "MHz", 79, "M");

            }

            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", CONECTIVIDAD_GILAT, 83, "L");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", CONECTIVIDAD_NODO_TERMINAL, 84, "L");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", CONECTIVIDAD_NODO_TERMINAL, 85, "L");

            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", NOMBRES_APELLIDOS_ENCARGADO, 95, "F");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", DOC_IDENTIDAD_ENCARGADO, 96, "F");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", CELULAR_CONTACTO_ENCARGADO, 97, "F");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", EMAIL_ENCARGADO_IIBB, 98, "F");

            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", NOMBRES_APELLIDOS_REPRESENTANTE, 119, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", DOC_IDENTIDAD_REPRESENTANTE, 119, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", CARGO_REPRESENTANTE_IIBB, 119, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", EMAIL_REPRESENTANTE_IIBB, 119, "L");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", NOMBRES_APELLIDOS_REPR_OPERADOR, 120, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", DOC_IDENTIDAD_REPR_OPERADOR, 120, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", CARGO_REPRESENTANTE_OPERADOR, 120, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Acta de Instalación FITEL", EMAIL_REPR_OPERADOR, 120, "L");


            #endregion

            #region Configuracion y Pruebas

            ExcelToolsBL.UpdateCell(excelGenerado, "2 Configuración y Pruebas", NOMBRE_IIBB, 10, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "2 Configuración y Pruebas", CODIGO_IIBB, 13, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "2 Configuración y Pruebas", NOMBRE_IIBB, 14, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "2 Configuración y Pruebas", TIPO_INSTITUCION, 15, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "2 Configuración y Pruebas", LATITUD, 18, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "2 Configuración y Pruebas", LONGITUD, 19, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "2 Configuración y Pruebas", CODIGO_NODO, 23, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "2 Configuración y Pruebas", IP_IIBB, 27, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "2 Configuración y Pruebas", NOMBRES_APELLIDOS_REPRESENTANTE, 33, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "2 Configuración y Pruebas", CARGO_REPRESENTANTE_IIBB, 34, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "2 Configuración y Pruebas", CELULAR_CONTACTO_REPRESENTANTE, 35, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "2 Configuración y Pruebas", ODU_CPE, 38, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "2 Configuración y Pruebas", ALTURA_MASTIL, 41, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "2 Configuración y Pruebas", TIPO_MASTIL, 42, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "2 Configuración y Pruebas", ALTITUDmsnm, 62, "D");


            #endregion

            #region Pantallas de Configuracion 

             ExcelToolsBL.AddImageDocument(false, excelGenerado, "3 Pantallas de Configuración", mPANT_CONF_ACCESS_POINt, "", 11, 3, 907, 491);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "3 Pantallas de Configuración", mPANT_CONF_ROUTER, "", 54, 3, 908, 592);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "3 Pantallas de Configuración", mPANT_CONF_SWITCH01, "", 101, 4, 768, 401);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "3 Pantallas de Configuración", mPANT_CONF_SWITCH02, "", 121, 4, 732, 339);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "3 Pantallas de Configuración", mPANT_CONF_UPS, "", 138, 4, 767, 399);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "3 Pantallas de Configuración", mPANT_CONF_ALLINONE01, "", 166, 4, 746, 447);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "3 Pantallas de Configuración", mPANT_CONF_ALLINONE01, "", 169, 4, 736, 447);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "3 Pantallas de Configuración", mPANT_CONF_IMPRESORA, "", 196, 5, 593, 399);

            #endregion

            #region Medicion SPAT   
            ExcelToolsBL.UpdateCell(excelGenerado, "4 Medición SPAT", MSPT_MEDIDA1_VALORMEDIO, 7, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "4 Medición SPAT", MSPT_MEDIDA2_VALORMEDIO, 8, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "4 Medición SPAT", MSPT_MEDIDA3_VALORMEDIO, 9, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "4 Medición SPAT", MSPTP_MEDIDA1_VALORMEDIO, 14, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "4 Medición SPAT", MSPTP_MEDIDA2_VALORMEDIO, 15, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "4 Medición SPAT", MSPTP_MEDIDA1_VALORMEDIO, 16, "I");

            #endregion

            #region Material IIBB

            ExcelToolsBL.UpdateCell(excelGenerado, "6 Material IIBB CPE", "INSTITUCION  " + NOMBRE_IIBB, 12, "F");

            foreach (DataRow dr in dt2.Rows)
            {
                String EQUIPO = dr["EQUIPO"].ToString();
                String MARCA = dr["MARCA"].ToString();
                String MODELO = dr["MODELO"].ToString();
                String Nro_SERIE = dr["Nro_SERIE"].ToString();

                int ind = Convert.ToInt32(dt2.Rows.IndexOf(dr));

                ExcelToolsBL.UpdateCell(excelGenerado, "6 Material IIBB CPE", Convert.ToString(ind + 1), 17 + ind, "B");
                ExcelToolsBL.UpdateCell(excelGenerado, "6 Material IIBB CPE", EQUIPO, 17 + ind, "C");
                ExcelToolsBL.UpdateCell(excelGenerado, "6 Material IIBB CPE", "1", 17 + ind, "D");
                ExcelToolsBL.UpdateCell(excelGenerado, "6 Material IIBB CPE", MARCA, 17 + ind, "E");
                ExcelToolsBL.UpdateCell(excelGenerado, "6 Material IIBB CPE", MODELO, 17 + ind, "F");
                ExcelToolsBL.UpdateCell(excelGenerado, "6 Material IIBB CPE", Nro_SERIE, 17 + ind, "G");
            }

            foreach (DataRow dr in dt3.Rows)
            {
                String DESCRIPCION = dr["DESCRIPCION"].ToString();
                String CODIGO = dr["CODIGO"].ToString();
                String UNIDAD = dr["UNIDAD"].ToString();
                String CANTIDAD = dr["CANTIDAD"].ToString();

                int ind = Convert.ToInt32(dt3.Rows.IndexOf(dr));

                ExcelToolsBL.UpdateCell(excelGenerado, "6 Material IIBB CPE", Convert.ToString(ind + 1), 32 + ind, "B");
                ExcelToolsBL.UpdateCell(excelGenerado, "6 Material IIBB CPE", DESCRIPCION, 32 + ind, "C");
                ExcelToolsBL.UpdateCell(excelGenerado, "6 Material IIBB CPE", CODIGO, 32 + ind, "D");
                ExcelToolsBL.UpdateCell(excelGenerado, "6 Material IIBB CPE", UNIDAD, 32 + ind, "F");
                ExcelToolsBL.UpdateCell(excelGenerado, "6 Material IIBB CPE", CANTIDAD, 32 + ind, "G");

            }

            #endregion

            #region Reporte Fotografico

            ExcelToolsBL.UpdateCell(excelGenerado, "7 Reporte Fotográfico IIBB CPE", CODIGO_IIBB + NOMBRE_IIBB, 7, "B");
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO1_PAN_LOCALIDAD, "", 11,3, 904, 505);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO2_FACHADA_INSTITUCION, "", 47,3, 900,457);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO3_1_IMPRESORA, "", 82,3,404, 251);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO3_2_SWITCH, "", 82,14,401, 251);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO3_3_ROUTER, "", 101, 3, 404, 282);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO3_4_PC_ENCENDIDAS, "", 101,14,403, 283);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO3_5_PC_UPS, "", 120,3,406,423);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO3_6_ACCESS_POINT, "", 120, 14, 402, 425);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO4_1_ODU_CPE, "", 141, 3, 404,293);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO4_2_MASTIL, "", 141, 14, 404,291);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO4_3_PAN_ANT_INSTAL_MASTIL, "", 160, 3, 405, 302);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO4_4_RECORRIDO_SFTP_CATSE, "", 160, 14, 405,305);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO4_5_INGRESO_SFTP, "", 178, 3, 406, 283);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO4_6_RECORRIDO_SFTP_CANALETA, "", 178, 14, 403, 282);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO4_7_POE, "", 197, 3, 405, 282);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO4_8_PATCH_POE_ROUTER, "", 197, 14, 402,281);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO5_1_TABLERO_GENERAL_SECUNDARIO, "", 219, 3, 405, 470);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO5_2_INSTALACION_BREAKER, "", 219, 14, 402, 471);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO5_3_CABLE_CONEXION_ELECTRICA, "", 237,3, 404, 471);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO5_4_TOMAS_ENERGIA, "", 237, 14, 406, 470);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO5_5_FOTO_INTERNA_INST_BREAKER, "", 254, 14, 401, 461);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO6_1_DNI_DJREPRESENTANTE_ABONADO, "", 274, 3, 450,354);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO6_2_DNI_DJREPRESENTANTE_ABONADO, "", 274, 13, 447, 355);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO7_1_SWITCH, "", 294, 3, 405,280);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO7_2_ROUTER, "", 294, 14, 404,281);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO7_3_REGLETA_ENERGIA, "", 311, 3, 405, 327);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO7_4_UPS, "",311, 14, 401, 325);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO7_5_COMPUTADORAS, "", 328, 3, 405, 285);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO7_6_ACESS_POINT, "", 328, 14, 402, 286);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO7_7_IMPRESORA, "", 345, 3, 405, 285);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO7_8_PAN_SALA_EQUIPOS, "", 345, 14, 405, 285);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO7_9_JACK_RJ45, "", 361, 3, 405, 285);

            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO8_1_INSTALACION_POZO_TIERRA, "", 381, 3, 406, 291);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO8_2_CONEX_CAJA_REGISTRO, "", 381, 14,402, 293);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO8_3_ESCALA_UTIL_RESULT_MEDICION1, "", 400, 3, 406, 376);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO8_4_ESCALA_UTIL_RESULT_MEDICION2, "", 400, 14, 403, 375);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO8_5_ESCALA_UTIL_RESULT_MEDICION3, "", 416, 3, 404, 379);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO9_1_INSTAL_POZO_TIERRA_1, "", 436, 3, 406, 294);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO9_2_INSTAL_POZO_TIERRA_2, "", 436, 14, 402, 295);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO9_3_ESCALA_UTIL_RESULT_MEDICION1, "", 455, 3, 404, 375);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO9_4_ESCALA_UTIL_RESULT_MEDICION2, "", 455, 14, 404, 377);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO9_5_ESCALA_UTIL_RESULT_MEDICION3, "", 471, 3, 404, 374);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO10_1_PANT_CONF_HOME, "", 492, 3, 406, 286);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO10_2_PANT_CONF_SECURITY, "", 492, 14, 402, 286);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO10_3_PANT_CONF_RADIO_1, "", 509, 3, 405, 285);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO10_4_PANT_CONF_RADIO_2, "", 509, 14, 402, 284);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO10_5_CONF_SISTEMA_1, "", 526, 3, 407, 285);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO10_6_CONF_SISTEMA_2, "", 526, 14, 402, 285);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO10_7_PANT_CONF_NETWORK_1, "", 543, 3, 405, 285);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO10_8_PANT_CONF_NETWORK_2, "", 543, 14, 405, 285);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO10_9_PANT_CONF_MONITOR_WIRELESS, "", 560, 3, 405, 285);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO10_10_CONF_SISTEMA_TOOLS, "", 560, 14, 405, 285);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO11_1_MON_CONEX_SITIO_WEB, "", 580, 3, 895, 479);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO11_2_MON_CONECTIVIDAD_NODO_TERMINAL, "", 596, 3, 895, 479);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "7 Reporte Fotográfico IIBB CPE", mFOTO11_3_MON_CONECTIVIDAD_NODO_DISTRITAL, "", 612,3, 895, 479);
            #endregion

            #region Serie equipos fotos

            ExcelToolsBL.AddImageDocument(false, excelGenerado, "8 Serie Equipos (fotos)", mFOTO_EPMP_1000_FORCE_180, "", 15,2, 1668, 393);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "8 Serie Equipos (fotos)", mFOTO_1_ACCESS_POINT_SERIE, "", 34, 2, 709, 627);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "8 Serie Equipos (fotos)", mFOTO_2_SWITCH_SERIE, "", 34, 13, 719, 626);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "8 Serie Equipos (fotos)", mFOTO_3_ROUTER_SERIE, "", 54, 2, 710, 533);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "8 Serie Equipos (fotos)", mFOTO_4_IMPRESORA_SERIE, "", 54, 13, 719, 533);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "8 Serie Equipos (fotos)", mFOTO_5_UPS_SERIE, "", 75, 8, 612, 560);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "8 Serie Equipos (fotos)", mFOTO_6_PC01_SERIE, "", 98, 8, 611, 556);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "8 Serie Equipos (fotos)", mFOTO_7_PC02_SERIE, "", 121,2, 709, 549);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "8 Serie Equipos (fotos)", mFOTO_8_PC03_SERIE, "", 122, 13, 719, 547);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "8 Serie Equipos (fotos)", mFOTO_9_PC04_SERIE, "", 141, 2, 710, 534);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "8 Serie Equipos (fotos)", mFOTO_10_PC05_SERIE, "", 141, 13, 720, 534);


            #endregion

            #endregion

            #region Ruta del Zip
            String rutaNodo = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo;
            if (Directory.Exists(rutaNodo))
            {
                String rutaParcial = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\IE01\\";
                if (!Directory.Exists(rutaParcial)) Directory.CreateDirectory(rutaParcial);
                String rutaAlterna = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\IE01\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";
                File.Copy(excelGenerado, rutaAlterna, true);
            }
            #endregion
        }


        public void ActaSeguridadAcceso(String IdNodo, String IdTarea, String valorCadena1, String rutaPlantilla)
        {
            DBBaseDatos baseDatosDA = new DBBaseDatos();

            baseDatosDA.Configurar();
            baseDatosDA.Conectar();
            DataTable dt = new DataTable();
            DataTable dt1 = new DataTable();
            DataTable dt2 = new DataTable();

            try
            {
                baseDatosDA.CrearComando("USP_R_ACTA_SEGURIDAD__ACCESO", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                dt = baseDatosDA.EjecutarConsultaDataTable();

                baseDatosDA.CrearComando("USP_R_EQUIPAMIENTOS_SEGURIDAD_ACCESO", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                dt1 = baseDatosDA.EjecutarConsultaDataTable();

                baseDatosDA.CrearComando("USP_R_MATERIALES_SEGURIDAD_ACCESO", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                dt2 = baseDatosDA.EjecutarConsultaDataTable();



              
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
            #region valores_String
            String NOMBRE_NODO = "NODO " + dt.Rows[0]["NOMBRE_NODO"].ToString();
            String CODIGO_NODO = dt.Rows[0]["CODIGO_NODO"].ToString();
            String TIPO_NODO = dt.Rows[0]["TIPO_NODO"].ToString();

            String fechaSQL = dt.Rows[0]["FECHA"].ToString();
            String FECHA = "";
            if (fechaSQL != "")
            {
                DateTime dtFecha = DateTime.Parse(fechaSQL);
                FECHA = dtFecha.ToString("dd/MM/yyyy");
            }
            else { FECHA = ""; }

            //String POWER_CABLE_3X14AWG = ds.Tables[0].Rows[0]["POWER_CABLE_3X14AWG"].ToString();
            //String OUTDOOR_CABLE_2X0_22SQMM_RED_BLACK = ds.Tables[0].Rows[0]["OUTDOOR_CABLE_2X0_22SQMM_RED_BLACK"].ToString();
            //String OUTDOOR_CABLE_4X0_22SQMM = ds.Tables[0].Rows[0]["OUTDOOR_CABLE_4X0_22SQMM"].ToString();
            //String SILICONA_TRANSPARENTE_200ML = ds.Tables[0].Rows[0]["SILICONA_TRANSPARENTE_200ML"].ToString();
            //String TUBO_CORRUGADO_PLEGABLE_PVC_20MM = ds.Tables[0].Rows[0]["TUBO_CORRUGADO_PLEGABLE_PVC_20MM"].ToString();
            //String SPIRAL_WRAP_12MM_WHITE = ds.Tables[0].Rows[0]["SPIRAL_WRAP_12MM_WHITE"].ToString();
            //String STEEL_FLEXIBLE_CONDUIT_34_DFX_LT = ds.Tables[0].Rows[0]["STEEL_FLEXIBLE_CONDUIT_34_DFX_LT"].ToString();
            //String GROUND_CABLE_AWG_10_YELLOWGREEN = ds.Tables[0].Rows[0]["GROUND_CABLE_AWG_10_YELLOWGREEN"].ToString();
            //String DATA_CABLE_CAT5E_FOR_OUTDOOR = ds.Tables[0].Rows[0]["DATA_CABLE_CAT5E_FOR_OUTDOOR"].ToString();
            //String LAN_CABLE_CAT5E_UTP_24AWG_LSZH_GREY = ds.Tables[0].Rows[0]["LAN_CABLE_CAT5E_UTP_24AWG_LSZH_GREY"].ToString();
            //String PVC_TAPE_25M_X_19MM_BLACK = ds.Tables[0].Rows[0]["PVC_TAPE_25M_X_19MM_BLACK"].ToString();

            String fechaSQL_1 = dt.Rows[0]["EXTINGUIDOR_EXT_FECHA_EXPIRACION"].ToString();
            String EXTINGUIDOR_EXT_FECHA_EXPIRACION = "";
            if (fechaSQL_1 != "")
            {
                DateTime dtFecha = DateTime.Parse(fechaSQL_1);
                EXTINGUIDOR_EXT_FECHA_EXPIRACION = dtFecha.ToString("dd/MM/yyyy");
            }
            else { EXTINGUIDOR_EXT_FECHA_EXPIRACION = ""; }

            String fechaSQL_2 = dt.Rows[0]["EXTINGUIDOR_INT_FECHA_EXPIRACION"].ToString();
            String EXTINGUIDOR_INT_FECHA_EXPIRACION = "";
            if (fechaSQL_2 != "")
            {
                DateTime dtFecha = DateTime.Parse(fechaSQL_2);
                EXTINGUIDOR_INT_FECHA_EXPIRACION = dtFecha.ToString("dd/MM/yyyy");
            }
            else { EXTINGUIDOR_INT_FECHA_EXPIRACION = ""; }
            String DEPARTAMENTO = dt.Rows[0]["DEPARTAMENTO"].ToString();
            String SERIAL_CONTROLADOR = dt.Rows[0]["SERIAL_CONTROLADOR"].ToString();
            //IP CONTROLADOR
            String IP_CONTROLADOR = dt.Rows[0]["IP_CONTROLADOR"].ToString();

            #endregion

            #region valores binarios
            byte[] FACHADA_DEL_NODO = (byte[])dt.Rows[0]["FACHADA_DEL_NODO"];
            MemoryStream mFACHADA_DEL_NODO = new MemoryStream(FACHADA_DEL_NODO);
            byte[] SALA_EQUIPOS_PANORAMICA_RACK = (byte[])dt.Rows[0]["SALA_EQUIPOS_PANORAMICA_RACK"];
            MemoryStream mSALA_EQUIPOS_PANORAMICA_RACK = new MemoryStream(SALA_EQUIPOS_PANORAMICA_RACK);
            byte[] PANORAMICA_INTERIOR_01 = (byte[])dt.Rows[0]["PANORAMICA_INTERIOR_01"];
            MemoryStream mPANORAMICA_INTERIOR_01 = new MemoryStream(PANORAMICA_INTERIOR_01);
            byte[] PANORAMICA_INTERIOR_02 = (byte[])dt.Rows[0]["PANORAMICA_INTERIOR_02"];
            MemoryStream mPANORAMICA_INTERIOR_02 = new MemoryStream(PANORAMICA_INTERIOR_02);
            byte[] PANORAMICA_EQUIPOS_PATIO = (byte[])dt.Rows[0]["PANORAMICA_EQUIPOS_PATIO"];
            MemoryStream mPANORAMICA_EQUIPOS_PATIO = new MemoryStream(PANORAMICA_EQUIPOS_PATIO);
            byte[] BREAKER_ASIGNADO_PARA_SEGURIDAD = (byte[])dt.Rows[0]["BREAKER_ASIGNADO_PARA_SEGURIDAD"];
            MemoryStream mBREAKER_ASIGNADO_PARA_SEGURIDAD = new MemoryStream(BREAKER_ASIGNADO_PARA_SEGURIDAD);
            byte[] CERRADURA_ELECTROMAGNETICA_EXTERNA = (byte[])dt.Rows[0]["CERRADURA_ELECTROMAGNETICA_EXTERNA"];
            MemoryStream mCERRADURA_ELECTROMAGNETICA_EXTERNA = new MemoryStream(CERRADURA_ELECTROMAGNETICA_EXTERNA);
            byte[] CERRADURA_ELECTROMAGNETICA_EXTERNA2 = (byte[])dt.Rows[0]["CERRADURA_ELECTROMAGNETICA_EXTERNA2"];
            MemoryStream mCERRADURA_ELECTROMAGNETICA_EXTERNA2 = new MemoryStream(CERRADURA_ELECTROMAGNETICA_EXTERNA2);
            byte[] SENSOR_MAGNETICO_EXTERMO = (byte[])dt.Rows[0]["SENSOR_MAGNETICO_EXTERMO"];
            MemoryStream mSENSOR_MAGNETICO_EXTERMO = new MemoryStream(SENSOR_MAGNETICO_EXTERMO);
            byte[] SENSOR_MAGNETICO_EXTERNO2 = (byte[])dt.Rows[0]["SENSOR_MAGNETICO_EXTERNO2"];
            MemoryStream mSENSOR_MAGNETICO_EXTERNO2 = new MemoryStream(SENSOR_MAGNETICO_EXTERNO2);
            byte[] CERRADURA_ELECTRICA_EXTERNA = (byte[])dt.Rows[0]["CERRADURA_ELECTRICA_EXTERNA"];
            MemoryStream mCERRADURA_ELECTRICA_EXTERNA = new MemoryStream(CERRADURA_ELECTRICA_EXTERNA);
            byte[] SENSOR_MOVIMIENTO_90_EXTERNO_N1 = (byte[])dt.Rows[0]["SENSOR_MOVIMIENTO_90_EXTERNO_N1"];
            MemoryStream mSENSOR_MOVIMIENTO_90_EXTERNO_N1 = new MemoryStream(SENSOR_MOVIMIENTO_90_EXTERNO_N1);
            byte[] SENSOR_MOVIMIENTO_90_EXTERNO_N2 = (byte[])dt.Rows[0]["SENSOR_MOVIMIENTO_90_EXTERNO_N2"];
            MemoryStream mSENSOR_MOVIMIENTO_90_EXTERNO_N2 = new MemoryStream(SENSOR_MOVIMIENTO_90_EXTERNO_N2);
            byte[] SIRENA_ESTROBOSCOPICA = (byte[])dt.Rows[0]["SIRENA_ESTROBOSCOPICA"];
            MemoryStream mSIRENA_ESTROBOSCOPICA = new MemoryStream(SIRENA_ESTROBOSCOPICA);
            byte[] LECTOR_BIOMETRICO = (byte[])dt.Rows[0]["LECTOR_BIOMETRICO"];
            MemoryStream mLECTOR_BIOMETRICO = new MemoryStream(LECTOR_BIOMETRICO);
            byte[] LECTOR_TARJETA = (byte[])dt.Rows[0]["LECTOR_TARJETA"];
            MemoryStream mLECTOR_TARJETA = new MemoryStream(LECTOR_TARJETA);
            byte[] CAMARA_EXTERIOR_PTZ = (byte[])dt.Rows[0]["CAMARA_EXTERIOR_PTZ"];
            MemoryStream mCAMARA_EXTERIOR_PTZ = new MemoryStream(CAMARA_EXTERIOR_PTZ);
            byte[] EXTINTOR_EXTERIOR = (byte[])dt.Rows[0]["EXTINTOR_EXTERIOR"];
            MemoryStream mEXTINTOR_EXTERIOR = new MemoryStream(EXTINTOR_EXTERIOR);
            byte[] SENSOR_MAGNETICO_INTERNO = (byte[])dt.Rows[0]["SENSOR_MAGNETICO_INTERNO"];
            MemoryStream mSENSOR_MAGNETICO_INTERNO = new MemoryStream(SENSOR_MAGNETICO_INTERNO);
            byte[] SENSOR_MAGNETICO_INTERNO_2 = (byte[])dt.Rows[0]["SENSOR_MAGNETICO_INTERNO_2"];
            MemoryStream mSENSOR_MAGNETICO_INTERNO_2 = new MemoryStream(SENSOR_MAGNETICO_INTERNO_2);
            byte[] SENSOR_OCUPACIONAL = (byte[])dt.Rows[0]["SENSOR_OCUPACIONAL"];
            MemoryStream mSENSOR_OCUPACIONAL = new MemoryStream(SENSOR_OCUPACIONAL);
            byte[] SENSOR_DE_HUMO = (byte[])dt.Rows[0]["SENSOR_DE_HUMO"];
            MemoryStream mSENSOR_DE_HUMO = new MemoryStream(SENSOR_DE_HUMO);
            byte[] SENSOR_MOVIMIENTO_360 = (byte[])dt.Rows[0]["SENSOR_MOVIMIENTO_360"];
            MemoryStream mSENSOR_MOVIMIENTO_360 = new MemoryStream(SENSOR_MOVIMIENTO_360);
            byte[] SENSOR_DE_INUNDACION = (byte[])dt.Rows[0]["SENSOR_DE_INUNDACION"];
            MemoryStream mSENSOR_DE_INUNDACION = new MemoryStream(SENSOR_DE_INUNDACION);
            byte[] CAMARA_PTZ_INTERIOR = (byte[])dt.Rows[0]["CAMARA_PTZ_INTERIOR"];
            MemoryStream mCAMARA_PTZ_INTERIOR = new MemoryStream(CAMARA_PTZ_INTERIOR);
            byte[] EXTINTOR_INTERIOR = (byte[])dt.Rows[0]["EXTINTOR_INTERIOR"];
            MemoryStream mEXTINTOR_INTERIOR = new MemoryStream(EXTINTOR_INTERIOR);
            byte[] RELE_EQUIPO_INTERO = (byte[])dt.Rows[0]["RELE_EQUIPO_INTERO"];
            MemoryStream mRELE_EQUIPO_INTERO = new MemoryStream(RELE_EQUIPO_INTERO);
            byte[] CONTROLADOR_NVR_SWITCH = (byte[])dt.Rows[0]["CONTROLADOR_NVR_SWITCH"];
            MemoryStream mCONTROLADOR_NVR_SWITCH = new MemoryStream(CONTROLADOR_NVR_SWITCH);
            byte[] ATERRAMIENTO_CONTROLADOR = (byte[])dt.Rows[0]["ATERRAMIENTO_CONTROLADOR"];
            MemoryStream mATERRAMIENTO_CONTROLADOR = new MemoryStream(ATERRAMIENTO_CONTROLADOR);
            byte[] ATERRAMIENTO_NVR_POE = (byte[])dt.Rows[0]["ATERRAMIENTO_NVR_POE"];
            MemoryStream mATERRAMIENTO_NVR_POE = new MemoryStream(ATERRAMIENTO_NVR_POE);
            byte[] ATERRAMIENTO_NVR_POE_2 = (byte[])dt.Rows[0]["ATERRAMIENTO_NVR_POE_2"];
            MemoryStream mATERRAMIENTO_NVR_POE_2 = new MemoryStream(ATERRAMIENTO_NVR_POE_2);
            byte[] ATERRAMIENTO_A_BARRA = (byte[])dt.Rows[0]["ATERRAMIENTO_A_BARRA"];
            MemoryStream mATERRAMIENTO_A_BARRA = new MemoryStream(ATERRAMIENTO_A_BARRA);
            byte[] SERIAL_NUMBER_SENSOR_MOVIMIENTO_1 = (byte[])dt.Rows[0]["SERIAL_NUMBER_SENSOR_MOVIMIENTO_1"];
            MemoryStream mSERIAL_NUMBER_SENSOR_MOVIMIENTO_1 = new MemoryStream(SERIAL_NUMBER_SENSOR_MOVIMIENTO_1);
            byte[] SERIAL_NUMBER_SENSOR_MOVIMIENTO_2 = (byte[])dt.Rows[0]["SERIAL_NUMBER_SENSOR_MOVIMIENTO_2"];
            MemoryStream mSERIAL_NUMBER_SENSOR_MOVIMIENTO_2 = new MemoryStream(SERIAL_NUMBER_SENSOR_MOVIMIENTO_2);
            byte[] SERIAL_NUMBER_SWITCH_POE_NVR = (byte[])dt.Rows[0]["SERIAL_NUMBER_SWITCH_POE_NVR"];
            MemoryStream mSERIAL_NUMBER_SWITCH_POE_NVR = new MemoryStream(SERIAL_NUMBER_SWITCH_POE_NVR);
            byte[] SERIAL_NUMBER_SWITCH_POE_NVR_2 = (byte[])dt.Rows[0]["SERIAL_NUMBER_SWITCH_POE_NVR_2"];
            MemoryStream mSERIAL_NUMBER_SWITCH_POE_NVR_2 = new MemoryStream(SERIAL_NUMBER_SWITCH_POE_NVR_2);
            byte[] SERIAL_NUMBER_CONTROLADOR = (byte[])dt.Rows[0]["SERIAL_NUMBER_CONTROLADOR"];
            MemoryStream mSERIAL_NUMBER_CONTROLADOR = new MemoryStream(SERIAL_NUMBER_CONTROLADOR);
            byte[] ETIQUETADOS_EQUIPOS_CONTROLADOR = (byte[])dt.Rows[0]["ETIQUETADOS_EQUIPOS_CONTROLADOR"];
            MemoryStream mETIQUETADOS_EQUIPOS_CONTROLADOR = new MemoryStream(ETIQUETADOS_EQUIPOS_CONTROLADOR);
            byte[] ETIQUETADOS_EQUIPOS_NVR = (byte[])dt.Rows[0]["ETIQUETADOS_EQUIPOS_NVR"];
            MemoryStream mETIQUETADOS_EQUIPOS_NVR = new MemoryStream(ETIQUETADOS_EQUIPOS_NVR);
            byte[] CHECKLIST = (byte[])dt.Rows[0]["CHECKLIST"];
            MemoryStream mCHECKLIST = new MemoryStream(CHECKLIST);
            byte[] CAMARA_EXTERIOR_MODO_NORMAL_POS1 = (byte[])dt.Rows[0]["CAMARA_EXTERIOR_MODO_NORMAL_POS1"];
            MemoryStream mCAMARA_EXTERIOR_MODO_NORMAL_POS1 = new MemoryStream(CAMARA_EXTERIOR_MODO_NORMAL_POS1);
            byte[] CAMARA_EXTERIOR_MODO_NORMAL_POS2 = (byte[])dt.Rows[0]["CAMARA_EXTERIOR_MODO_NORMAL_POS2"];
            MemoryStream mCAMARA_EXTERIOR_MODO_NORMAL_POS2 = new MemoryStream(CAMARA_EXTERIOR_MODO_NORMAL_POS2);
            byte[] CAMARA_INTERIOR_MODO_NORMAL = (byte[])dt.Rows[0]["CAMARA_INTERIOR_MODO_NORMAL"];
            MemoryStream mCAMARA_INTERIOR_MODO_NORMAL = new MemoryStream(CAMARA_INTERIOR_MODO_NORMAL);
            byte[] CAMARA_INTERIOR_MODO_INFRARROJO = (byte[])dt.Rows[0]["CAMARA_INTERIOR_MODO_INFRARROJO"];
            MemoryStream mCAMARA_INTERIOR_MODO_INFRARROJO = new MemoryStream(CAMARA_INTERIOR_MODO_INFRARROJO);
            byte[] TPA_PUERTA_PRINCIPAL_ABIERTA = (byte[])dt.Rows[0]["TPA_PUERTA_PRINCIPAL_ABIERTA"];
            MemoryStream mTPA_PUERTA_PRINCIPAL_ABIERTA = new MemoryStream(TPA_PUERTA_PRINCIPAL_ABIERTA);
            byte[] TPA_PUERTA_SALAS_EQUIPOS_ABIERTA = (byte[])dt.Rows[0]["TPA_PUERTA_SALAS_EQUIPOS_ABIERTA"];
            MemoryStream mTPA_PUERTA_SALAS_EQUIPOS_ABIERTA = new MemoryStream(TPA_PUERTA_SALAS_EQUIPOS_ABIERTA);
            byte[] TPA_CAMARA_INTERNA = (byte[])dt.Rows[0]["TPA_CAMARA_INTERNA"];
            MemoryStream mTPA_CAMARA_INTERNA = new MemoryStream(TPA_CAMARA_INTERNA);
            byte[] TPA_CAMARA_EXTERNA = (byte[])dt.Rows[0]["TPA_CAMARA_EXTERNA"];
            MemoryStream mTPA_CAMARA_EXTERNA = new MemoryStream(TPA_CAMARA_EXTERNA);
            byte[] TPA_SENSOR_DE_ANIEGO = (byte[])dt.Rows[0]["TPA_SENSOR_DE_ANIEGO"];
            MemoryStream mTPA_SENSOR_DE_ANIEGO = new MemoryStream(TPA_SENSOR_DE_ANIEGO);
            byte[] TPA_SENSOR_DE_HUMO = (byte[])dt.Rows[0]["TPA_SENSOR_DE_HUMO"];
            MemoryStream mTPA_SENSOR_DE_HUMO = new MemoryStream(TPA_SENSOR_DE_HUMO);
            byte[] TPA_TAMPER_SENSOR_90_1 = (byte[])dt.Rows[0]["TPA_TAMPER_SENSOR_90_1"];
            MemoryStream mTPA_TAMPER_SENSOR_90_1 = new MemoryStream(TPA_TAMPER_SENSOR_90_1);
            byte[] TPA_MOVIMIENTO_SENSOR_90_1 = (byte[])dt.Rows[0]["TPA_MOVIMIENTO_SENSOR_90_1"];
            MemoryStream mTPA_MOVIMIENTO_SENSOR_90_1 = new MemoryStream(TPA_MOVIMIENTO_SENSOR_90_1);
            byte[] TPA_MASKING_SENSOR_90_1 = (byte[])dt.Rows[0]["TPA_MASKING_SENSOR_90_1"];
            MemoryStream mTPA_MASKING_SENSOR_90_1 = new MemoryStream(TPA_MASKING_SENSOR_90_1);
            byte[] TPA_TAMPER_SENSOR_90_2 = (byte[])dt.Rows[0]["TPA_TAMPER_SENSOR_90_2"];
            MemoryStream mTPA_TAMPER_SENSOR_90_2 = new MemoryStream(TPA_TAMPER_SENSOR_90_2);
            byte[] TPA_MOVIMIENTO_SENSOR_90_2 = (byte[])dt.Rows[0]["TPA_MOVIMIENTO_SENSOR_90_2"];
            MemoryStream mTPA_MOVIMIENTO_SENSOR_90_2 = new MemoryStream(TPA_MOVIMIENTO_SENSOR_90_2);
            byte[] TPA_MASKING_SENSOR_90_2 = (byte[])dt.Rows[0]["TPA_MASKING_SENSOR_90_2"];
            MemoryStream mTPA_MASKING_SENSOR_90_2 = new MemoryStream(TPA_MASKING_SENSOR_90_2);
            byte[] TPA_ALARMA_TAMPER_SENSOR_360 = (byte[])dt.Rows[0]["TPA_ALARMA_TAMPER_SENSOR_360"];
            MemoryStream mTPA_ALARMA_TAMPER_SENSOR_360 = new MemoryStream(TPA_ALARMA_TAMPER_SENSOR_360);
            byte[] TPA_ALARMA_MOVIMIENTO_SENSOR_360 = (byte[])dt.Rows[0]["TPA_ALARMA_MOVIMIENTO_SENSOR_360"];
            MemoryStream mTPA_ALARMA_MOVIMIENTO_SENSOR_360 = new MemoryStream(TPA_ALARMA_MOVIMIENTO_SENSOR_360);
            byte[] PING_CAMARA_1_INDOOR = (byte[])dt.Rows[0]["PING_CAMARA_1_INDOOR"];
            MemoryStream mPING_CAMARA_1_INDOOR = new MemoryStream(PING_CAMARA_1_INDOOR);
            byte[] PING_CAMARA_2_OUTDOOR = (byte[])dt.Rows[0]["PING_CAMARA_2_OUTDOOR"];
            MemoryStream mPING_CAMARA_2_OUTDOOR = new MemoryStream(PING_CAMARA_2_OUTDOOR);
            byte[] PING_CONTROLADOR = (byte[])dt.Rows[0]["PING_CONTROLADOR"];
            MemoryStream mPING_CONTROLADOR = new MemoryStream(PING_CONTROLADOR);
            byte[] PING_GATEWAY = (byte[])dt.Rows[0]["PING_GATEWAY"];
            MemoryStream mPING_GATEWAY = new MemoryStream(PING_GATEWAY);
            byte[] PING_NVR = (byte[])dt.Rows[0]["PING_NVR"];
            MemoryStream mPING_NVR = new MemoryStream(PING_NVR);
            byte[] PING_BIOMETRICO = (byte[])dt.Rows[0]["PING_BIOMETRICO"];
            MemoryStream mPING_BIOMETRICO = new MemoryStream(PING_BIOMETRICO);
            #endregion

            String excelGenerado = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";
            File.Copy(rutaPlantilla, excelGenerado, true);

            #region Agregando datos

            #region Caratula 
            ExcelToolsBL.UpdateCell(excelGenerado, "Carátula", "NODO  " + NOMBRE_NODO, 15, "B");
            ExcelToolsBL.UpdateCell(excelGenerado, "Carátula", NOMBRE_NODO + " - PERU", 17, "B");
            ExcelToolsBL.UpdateCell(excelGenerado, "Carátula", TIPO_NODO, 22, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "Carátula", CODIGO_NODO, 24, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "Carátula", FECHA, 26, "D");
            #endregion

            #region Acta Instalacion Aceptacion
            ExcelToolsBL.UpdateCell(excelGenerado, "Acta de Instal- aceptación", "NODO  " + NOMBRE_NODO, 15, "E");
            #endregion

            #region Reporte Fotografico

            ExcelToolsBL.UpdateCell(excelGenerado, "Reporte fotográfico", "SISTEMAS DE SEGURIDAD NODO  " + TIPO_NODO + " _" + NOMBRE_NODO, 5, "B");
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mFACHADA_DEL_NODO, "", 10, 3, 405, 260);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mSALA_EQUIPOS_PANORAMICA_RACK, "", 10, 14, 441, 263);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mPANORAMICA_INTERIOR_01, "", 26, 3, 407, 261);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mPANORAMICA_INTERIOR_02, "", 26, 14, 441, 260);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mPANORAMICA_EQUIPOS_PATIO, "", 42, 3, 405, 260);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mBREAKER_ASIGNADO_PARA_SEGURIDAD, "", 42, 14, 441, 261);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mCERRADURA_ELECTROMAGNETICA_EXTERNA, "", 59, 3, 182, 241);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mCERRADURA_ELECTROMAGNETICA_EXTERNA2, "", 59, 7, 222, 240);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mSENSOR_MAGNETICO_EXTERMO, "", 59, 14, 225, 239);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mSENSOR_MAGNETICO_EXTERNO2, "", 59, 19, 216, 240);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mCERRADURA_ELECTRICA_EXTERNA, "", 78, 3, 405, 308);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mSENSOR_MOVIMIENTO_90_EXTERNO_N1, "", 78, 14, 442, 306);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mSENSOR_MOVIMIENTO_90_EXTERNO_N2, "", 96, 3, 404, 267);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mSIRENA_ESTROBOSCOPICA, "", 96, 14, 440, 268);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mLECTOR_BIOMETRICO, "", 116, 3, 404, 300);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mLECTOR_TARJETA, "", 116, 14, 441, 295);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mCAMARA_EXTERIOR_PTZ, "", 135, 3, 404, 309);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mEXTINTOR_EXTERIOR, "", 135, 14, 443, 309);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mSENSOR_MAGNETICO_INTERNO, "", 156, 3, 181, 279);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mSENSOR_MAGNETICO_INTERNO_2, "", 156, 7, 225, 278);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mSENSOR_OCUPACIONAL, "", 156, 14, 443, 281);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mSENSOR_DE_HUMO, "", 172, 3, 405, 311);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mSENSOR_MOVIMIENTO_360, "", 172, 14, 442, 311);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mSENSOR_DE_INUNDACION, "", 188, 3, 404, 312);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mCAMARA_PTZ_INTERIOR, "", 188, 14, 441, 311);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mEXTINTOR_INTERIOR, "", 205, 3, 406, 461);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mRELE_EQUIPO_INTERO, "", 205, 14, 441, 464);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mCONTROLADOR_NVR_SWITCH, "", 222, 3, 889, 355);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mATERRAMIENTO_CONTROLADOR, "", 247, 3, 404, 303);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mATERRAMIENTO_NVR_POE, "", 247, 14, 226, 305);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mATERRAMIENTO_NVR_POE_2, "", 247, 19, 216, 304);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mATERRAMIENTO_A_BARRA, "", 263, 3,407,272);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mSERIAL_NUMBER_SENSOR_MOVIMIENTO_1, "", 282,3,407,278);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mSERIAL_NUMBER_SENSOR_MOVIMIENTO_2, "", 282,14,442,279);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mSERIAL_NUMBER_SWITCH_POE_NVR, "", 298,3,181,331);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mSERIAL_NUMBER_SWITCH_POE_NVR_2, "", 298, 7,226, 332);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mSERIAL_NUMBER_CONTROLADOR, "", 298,14,442,331);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mETIQUETADOS_EQUIPOS_CONTROLADOR, "", 317,5,677, 312);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mETIQUETADOS_EQUIPOS_NVR, "",335,5,678,311);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mCHECKLIST, "", 353,6,630,700);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mCAMARA_EXTERIOR_MODO_NORMAL_POS1, "",391,5,723,369);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mCAMARA_EXTERIOR_MODO_NORMAL_POS2,"",408,5,723,312);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mCAMARA_INTERIOR_MODO_NORMAL, "", 425,5,722,312);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mCAMARA_INTERIOR_MODO_INFRARROJO, "",442,5,679,341);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mTPA_PUERTA_PRINCIPAL_ABIERTA, "", 459,5,724,377);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mTPA_PUERTA_SALAS_EQUIPOS_ABIERTA, "",476,5,723,378);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mTPA_CAMARA_INTERNA,"",492,5,723,379);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mTPA_CAMARA_EXTERNA, "",507,5,722,378);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mTPA_SENSOR_DE_ANIEGO, "", 522,5,723,381);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mTPA_SENSOR_DE_HUMO, "", 539,5,724,379);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mTPA_TAMPER_SENSOR_90_1, "",554,5,722,371);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mTPA_MOVIMIENTO_SENSOR_90_1, "",569,5,724,372);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mTPA_MASKING_SENSOR_90_1,"",584,5,723,373);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mTPA_TAMPER_SENSOR_90_2, "",599,5,724,373);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mTPA_MOVIMIENTO_SENSOR_90_2,"",614,5,723,372);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mTPA_MASKING_SENSOR_90_2, "",629,5,722,374);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mTPA_ALARMA_TAMPER_SENSOR_360,"",644,5,722,372);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mTPA_ALARMA_MOVIMIENTO_SENSOR_360,"",659,5,722,372);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mPING_CAMARA_1_INDOOR,"",676,5,722,311);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mPING_CAMARA_2_OUTDOOR,"",693,5,723,309);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mPING_CONTROLADOR,"",710,5,723,312);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mPING_GATEWAY,"",727,5,724,360);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mPING_NVR, "",744,5,722,336);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte fotográfico", mPING_BIOMETRICO,"",761,5,723,338);
            #endregion

            #region Materiales

            ExcelToolsBL.UpdateCell(excelGenerado, "Materiales", "NODO  " + TIPO_NODO, 11, "C");
            ExcelToolsBL.UpdateCell(excelGenerado, "Materiales", CODIGO_NODO, 11, "F");


            foreach (DataRow dr in dt1.Rows)
            {
                String EQUIPO = dr["EQUIPO"].ToString();
                String MARCA = dr["MARCA"].ToString();
                String MODELO = dr["MODELO"].ToString();
                String Nro_SERIE = dr["Nro_SERIE"].ToString();

                int ind = Convert.ToInt32(dt1.Rows.IndexOf(dr));

                ExcelToolsBL.UpdateCell(excelGenerado, "Materiales", Convert.ToString(ind + 1), 16 + ind, "B");
                ExcelToolsBL.UpdateCell(excelGenerado, "Materiales", EQUIPO, 16 + ind, "C");
                ExcelToolsBL.UpdateCell(excelGenerado, "Materiales", "1", 16 + ind, "D");
                ExcelToolsBL.UpdateCell(excelGenerado, "Materiales", MARCA, 16 + ind, "E");
                ExcelToolsBL.UpdateCell(excelGenerado, "Materiales", MODELO, 16 + ind, "F");
                ExcelToolsBL.UpdateCell(excelGenerado, "Materiales", Nro_SERIE, 16 + ind, "G");
            }


            foreach (DataRow dr in dt2.Rows)
            {
                String DESCRIPCION = dr["DESCRIPCION"].ToString();
                String CODIGO = dr["CODIGO"].ToString();
                String UNIDAD = dr["UNIDAD"].ToString();
                String CANTIDAD = dr["CANTIDAD"].ToString();

                int ind = Convert.ToInt32(dt2.Rows.IndexOf(dr));

                ExcelToolsBL.UpdateCell(excelGenerado, "Materiales", Convert.ToString(ind + 1), 40 + ind, "B");
                ExcelToolsBL.UpdateCell(excelGenerado, "Materiales", DESCRIPCION, 40 + ind, "C");
                ExcelToolsBL.UpdateCell(excelGenerado, "Materiales", UNIDAD, 40 + ind, "E");
                ExcelToolsBL.UpdateCell(excelGenerado, "Materiales", CANTIDAD, 40 + ind, "F");

            }

            #endregion

            #region ATP
            ExcelToolsBL.UpdateCell(excelGenerado, "ATP", DEPARTAMENTO,9,"C");
            ExcelToolsBL.UpdateCell(excelGenerado, "ATP", FECHA, 6, "C");
            ExcelToolsBL.UpdateCell(excelGenerado, "ATP", CODIGO_NODO,9,"J");
            ExcelToolsBL.UpdateCell(excelGenerado, "ATP", NOMBRE_NODO,10,"J");
            ExcelToolsBL.UpdateCell(excelGenerado, "ATP", EXTINGUIDOR_EXT_FECHA_EXPIRACION, 43, "G");
            ExcelToolsBL.UpdateCell(excelGenerado, "ATP", EXTINGUIDOR_INT_FECHA_EXPIRACION, 44, "G");
            ExcelToolsBL.UpdateCell(excelGenerado, "ATP", SERIAL_CONTROLADOR, 8, "C");
            //IP CONTROLADOR
            ExcelToolsBL.UpdateCell(excelGenerado, "ATP", IP_CONTROLADOR, 13, "C");
            #endregion

            #endregion

            #region Ruta del Zip
            String rutaNodo = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo;
            if (Directory.Exists(rutaNodo))
            {
                String rutaParcial = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\ACCESO\\SEGURIDAD DISTRITAL\\";
                if (!Directory.Exists(rutaParcial)) Directory.CreateDirectory(rutaParcial);
                String rutaAlterna = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\ACCESO\\SEGURIDAD DISTRITAL\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";
                File.Copy(excelGenerado, rutaAlterna, true);
            }
            #endregion
        }

        public void Anexo2InventarioPTP(String IdNodo, String IdTarea, String valorCadena1, String rutaPlantilla)
        {
            DBBaseDatos baseDatosDA = new DBBaseDatos();

            baseDatosDA.Configurar();
            baseDatosDA.Conectar();
            DataTable dt = new DataTable();

            try
            {
                baseDatosDA.CrearComando("USP_R_ANEXO_02_INVENTARIO_PTP", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                dt = baseDatosDA.EjecutarConsultaDataTable();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            finally
            {

                baseDatosDA.Desconectar();
                baseDatosDA = null;
            }

            #region Valores
            String COD_NODO_A = dt.Rows[0]["COD_NODO_A"].ToString();
            byte[] SERIE_ANTENA_EST_A = (byte[])dt.Rows[0]["SERIE_ANTENA_EST_A"];
            MemoryStream mSERIE_ANTENA_EST_A = new MemoryStream(SERIE_ANTENA_EST_A);
            byte[] SERIE_ODU_EST_A = (byte[])dt.Rows[0]["SERIE_ODU_EST_A"];
            MemoryStream mSERIE_ODU_EST_A = new MemoryStream(SERIE_ODU_EST_A);
            byte[] SERIE_POE_EST_A = (byte[])dt.Rows[0]["SERIE_POE_EST_A"];
            MemoryStream mSERIE_POE_EST_A = new MemoryStream(SERIE_POE_EST_A);

            String COD_NODO_B = dt.Rows[0]["COD_NODO_B"].ToString();
            byte[] SERIE_ANTENA_EST_B = (byte[])dt.Rows[0]["SERIE_ANTENA_EST_B"];
            MemoryStream mSERIE_ANTENA_EST_B = new MemoryStream(SERIE_ANTENA_EST_B);
            byte[] SERIE_ODU_EST_B = (byte[])dt.Rows[0]["SERIE_ODU_EST_B"];
            MemoryStream mSERIE_ODU_EST_B = new MemoryStream(SERIE_ODU_EST_B);
            byte[] SERIE_POE_EST_B = (byte[])dt.Rows[0]["SERIE_POE_EST_B"];
            MemoryStream mSERIE_POE_EST_B = new MemoryStream(SERIE_POE_EST_B);

            CMM4BE CMM4A = new CMM4BE();
            List<CMM4BE> lstCMM4A = new List<CMM4BE>();
            CMM4A.Nodo.IdNodo = COD_NODO_A;
            lstCMM4A = CMM4BL.ListarCMM4(CMM4A);

            CMM4BE CMM4B = new CMM4BE();
            List<CMM4BE> lstCMM4B = new List<CMM4BE>();
            CMM4A.Nodo.IdNodo = COD_NODO_B;
            lstCMM4B = CMM4BL.ListarCMM4(CMM4B);

            #endregion

            // String usuarioWindows = Environment.UserName;
            String excelGenerado = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";

            File.Copy(rutaPlantilla, excelGenerado, true);

            #region Ingresando Valores

            if (!lstCMM4A.Count.Equals(0))
            {
                byte[] SERIE_CMM4_EST_A = (byte[])dt.Rows[0]["SERIE_CMM4_EST_A"];
                MemoryStream mSERIE_CMM4_EST_A = new MemoryStream(SERIE_CMM4_EST_A);
                byte[] SERIE_UGPS_EST_A = (byte[])dt.Rows[0]["SERIE_UGPS_EST_A"];
                MemoryStream mSERIE_UGPS_EST_A = new MemoryStream(SERIE_UGPS_EST_A);
                byte[] SERIE_CONVERSOR_EST_A = (byte[])dt.Rows[0]["SERIE_CONVERSOR_EST_A"];
                MemoryStream mSERIE_CONVERSOR_EST_A = new MemoryStream(SERIE_CONVERSOR_EST_A);

                ExcelToolsBL.AddImageDocument(false, excelGenerado, "11 Serie Logística", mSERIE_CMM4_EST_A, "", 37, 2, 704, 277);
                ExcelToolsBL.AddImageDocument(false, excelGenerado, "11 Serie Logística", mSERIE_UGPS_EST_A, "", 45,2, 706, 290);
                ExcelToolsBL.AddImageDocument(false, excelGenerado, "11 Serie Logística", mSERIE_CONVERSOR_EST_A, "", 52,2, 706, 298);

            }

            if (!lstCMM4B.Count.Equals(0))
            {


                byte[] SERIE_CMM4_EST_B = (byte[])dt.Rows[0]["SERIE_CMM4_EST_B"];
                MemoryStream mSERIE_CMM4_EST_B = new MemoryStream(SERIE_CMM4_EST_B);
                byte[] SERIE_UGPS_EST_B = (byte[])dt.Rows[0]["SERIE_UGPS_EST_B"];
                MemoryStream mSERIE_UGPS_EST_B = new MemoryStream(SERIE_UGPS_EST_B);
                byte[] SERIE_CONVERSOR_EST_B = (byte[])dt.Rows[0]["SERIE_CONVERSOR_EST_B"];
                MemoryStream mSERIE_CONVERSOR_EST_B = new MemoryStream(SERIE_CONVERSOR_EST_B);

                ExcelToolsBL.AddImageDocument(false, excelGenerado, "11 Serie Logística", mSERIE_CMM4_EST_B, "", 82,2, 706, 298);
                ExcelToolsBL.AddImageDocument(false, excelGenerado, "11 Serie Logística", mSERIE_UGPS_EST_B, "", 90,2,706, 295);
                ExcelToolsBL.AddImageDocument(false, excelGenerado, "11 Serie Logística", mSERIE_CONVERSOR_EST_B, "", 96,2, 706, 312);


            }

            ExcelToolsBL.UpdateCell(excelGenerado, "11 Serie Logística", "ENLACE " + COD_NODO_A + " - " + COD_NODO_B, 11, "B");
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "11 Serie Logística", mSERIE_ANTENA_EST_A, "", 15, 2, 707, 305);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "11 Serie Logística", mSERIE_ODU_EST_A, "", 22, 2, 707, 336);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "11 Serie Logística", mSERIE_POE_EST_A, "", 30, 2, 706, 287);

            ExcelToolsBL.AddImageDocument(false, excelGenerado, "11 Serie Logística", mSERIE_ANTENA_EST_B, "", 60, 2, 707, 284);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "11 Serie Logística", mSERIE_ODU_EST_B, "", 67, 2, 708, 323);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "11 Serie Logística", mSERIE_POE_EST_B, "", 75, 2, 706, 297);




            #endregion

            #region Ruta del Zip
            String rutaNodo = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo;
            if (Directory.Exists(rutaNodo))
            {
                String rutaParcial = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\ACCESO\\PTP_HC-0XX1 (450i)\\";
                if (!Directory.Exists(rutaParcial)) Directory.CreateDirectory(rutaParcial);
                String rutaAlterna = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\ACCESO\\PTP_HC-0XX1 (450i)\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";
                File.Copy(excelGenerado, rutaAlterna, true);
            }
            #endregion
        }

        public void ActaInstalacionPTPLicenciado(String IdNodo, String IdTarea, String valorCadena1, String rutaPlantilla)
        {
            DBBaseDatos baseDatosDA = new DBBaseDatos();

            baseDatosDA.Configurar();
            baseDatosDA.Conectar();
            DataTable dt = new DataTable();
            DataTable dt1 = new DataTable();
            DataTable dt2 = new DataTable();
            DataTable dt3 = new DataTable();
            DataTable dt4 = new DataTable();
            DataTable dt5 = new DataTable();
            DataTable dt6 = new DataTable();

            try
            {
                baseDatosDA.CrearComando("USP_R_ACTA_INSTALACION_PTP_LIC", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA",IdTarea ,true);
                dt = baseDatosDA.EjecutarConsultaDataTable();

                baseDatosDA.CrearComando("USP_R_EQUIPAMIENTOS_PTP_LIC_A", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                dt1 = baseDatosDA.EjecutarConsultaDataTable();

                baseDatosDA.CrearComando("USP_R_MATERIALES_PTP_LIC_A", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                dt2 = baseDatosDA.EjecutarConsultaDataTable();

                baseDatosDA.CrearComando("USP_R_EQUIPAMIENTOS_PTP_LIC_B", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                dt3 = baseDatosDA.EjecutarConsultaDataTable();

                baseDatosDA.CrearComando("USP_R_MATERIALES_PTP_LIC_B", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                dt4 = baseDatosDA.EjecutarConsultaDataTable();

                baseDatosDA.CrearComando("USP_R_MEDICION_ENLACE_PTP_LIC_A", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                dt5 = baseDatosDA.EjecutarConsultaDataTable();

                baseDatosDA.CrearComando("USP_R_MEDICION_ENLACE_PTP_LIC_B", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                dt6 = baseDatosDA.EjecutarConsultaDataTable();

            }
            catch (Exception)
            {

                throw;
            }
            finally
            {

                baseDatosDA.Desconectar();
                baseDatosDA = null;
            }

            #region Valores

            #region Caratula

            //NOMBRE LOCALIDAD
            String NODO_A = dt.Rows[0]["NODO_A"].ToString();
            String NODO_B = dt.Rows[0]["NODO_B"].ToString();
            String NOMBRE_NODO_A = dt.Rows[0]["NOMBRE_NODO_A"].ToString();
            String NOMBRE_NODO_B = dt.Rows[0]["NOMBRE_NODO_B"].ToString();
            String TIPO_NODO_A = dt.Rows[0]["TIPO_NODO_A"].ToString();
            String TIPO_NODO_B = dt.Rows[0]["TIPO_NODO_B"].ToString();
            String FRECUENCIA = dt.Rows[0]["FRECUENCIA"].ToString();

            String fechaSQL = dt.Rows[0]["FECHA"].ToString();
            String FECHA = "";
            if (fechaSQL != "")
            {
                DateTime dtFecha = DateTime.Parse(fechaSQL);
                FECHA = dtFecha.ToString("dd/MM/yyyy");
            }
            else { FECHA = ""; }

            #endregion

            #region Configuracion y Mediciones

            String DIREC_ESTACION_A = dt.Rows[0]["DIREC_ESTACION_A"].ToString();
            String DIREC_ESTACION_B = dt.Rows[0]["DIREC_ESTACION_B"].ToString();
            String DISTRITO_A = dt.Rows[0]["DISTRITO_A"].ToString();
            String DISTRITO_B = dt.Rows[0]["DISTRITO_B"].ToString();
            String PROVINCIA_A = dt.Rows[0]["PROVINCIA_A"].ToString();
            String PROVINCIA_B = dt.Rows[0]["PROVINCIA_B"].ToString();
            String DEPARTAMENTO_A = dt.Rows[0]["DEPARTAMENTO_A"].ToString();
            String DEPARTAMENTO_B = dt.Rows[0]["DEPARTAMENTO_B"].ToString();
            String LATITUD_A = dt.Rows[0]["LATITUD_A"].ToString();
            String LATITUD_B = dt.Rows[0]["LATITUD_B"].ToString();
            String LONGITUD_A = dt.Rows[0]["LONGITUD_A"].ToString();
            String LONGITUD_B = dt.Rows[0]["LONGITUD_B"].ToString();
            String ALTURAmsnm_A = dt.Rows[0]["ALTURAmsnm_A"].ToString();
            String ALTURAmsnm_B = dt.Rows[0]["ALTURAmsnm_B"].ToString();
            String REF_UBIC_EST_A = dt.Rows[0]["REF_UBIC_EST_A"].ToString();
            String REF_UBIC_EST_B = dt.Rows[0]["REF_UBIC_EST_B"].ToString();
            String ALTURA_TORRE_A = dt.Rows[0]["ALTURA_TORRE_A"].ToString();
            String ALTURA_TORRE_B = dt.Rows[0]["ALTURA_TORRE_B"].ToString();
            String DISTANCIA_A_B = dt.Rows[0]["DISTANCIA_A_B"].ToString();
            String MODULACION = dt.Rows[0]["MODULACION"].ToString();
            String FRECUENCIA_TX_A = dt.Rows[0]["FRECUENCIA_TX_A"].ToString();
            String FRECUENCIA_TX_B = dt.Rows[0]["FRECUENCIA_TX_B"].ToString();
            String CANAL_ESTAC_A = dt.Rows[0]["CANAL_ESTAC_A"].ToString();
            String CANAL_ESTAC_B = dt.Rows[0]["CANAL_ESTAC_B"].ToString();
            String VELOCIDAD_HABILITADA = dt.Rows[0]["VELOCIDAD_HABILITADA"].ToString(); //ES LA MISMA PARA AMBOS NODOS
            String ANCHO_BANDA = dt.Rows[0]["ANCHO_BANDA"].ToString();
            String MODEL_ANTENA_A = dt.Rows[0]["MODEL_ANTENA_A"].ToString();
            String MODEL_ANTENA_B = dt.Rows[0]["MODEL_ANTENA_B"].ToString();
            String DIAMETRO_A = dt.Rows[0]["DIAMETRO_A"].ToString();
            String DIAMETRO_B = dt.Rows[0]["DIAMETRO_B"].ToString();
            String ALTURA_ANTENA_A = dt.Rows[0]["ALTURA_ANTENA_A"].ToString();
            String ALTURA_ANTENA_B = dt.Rows[0]["ALTURA_ANTENA_B"].ToString();
            String POLARIZACION_A = dt.Rows[0]["POLARIZACION_A"].ToString();
            String POLARIZACION_B = dt.Rows[0]["POLARIZACION_B"].ToString();
            String AZIMUT_A = dt.Rows[0]["AZIMUT_A"].ToString();
            String AZIMUT_B = dt.Rows[0]["AZIMUT_B"].ToString();
            String IP_NODO_A_1 = dt.Rows[0]["IP_NODO_A_1"].ToString();
            String IP_NODO_A_2 = dt.Rows[0]["IP_NODO_A_2"].ToString();
            String IP_NODO_A_3 = dt.Rows[0]["IP_NODO_A_3"].ToString();
            String IP_NODO_A_4 = dt.Rows[0]["IP_NODO_A_4"].ToString();
            String IP_NODO_B_1 = dt.Rows[0]["IP_NODO_B_1"].ToString();
            String IP_NODO_B_2 = dt.Rows[0]["IP_NODO_B_2"].ToString();
            String IP_NODO_B_3 = dt.Rows[0]["IP_NODO_B_3"].ToString();
            String IP_NODO_B_4 = dt.Rows[0]["IP_NODO_B_4"].ToString();
            String DEFAULT_GATE_AB_1 = dt.Rows[0]["DEFAULT_GATE_AB_1"].ToString();
            String DEFAULT_GATE_AB_2 = dt.Rows[0]["DEFAULT_GATE_AB_2"].ToString();
            String DEFAULT_GATE_AB_3 = dt.Rows[0]["DEFAULT_GATE_AB_3"].ToString();
            String DEFAULT_GATE_AB_4 = dt.Rows[0]["DEFAULT_GATE_AB_4"].ToString();
            String POTENCIA_A = dt.Rows[0]["POTENCIA_A"].ToString();
            String POTENCIA_B = dt.Rows[0]["POTENCIA_B"].ToString();
            String MARGEN_DES_A = dt.Rows[0]["MARGEN_DES_A"].ToString();
            String MARGEN_DES_B = dt.Rows[0]["MARGEN_DES_B"].ToString();
            String NIVEL_UMBRAL_A_B = dt.Rows[0]["NIVEL_UMBRAL_A_B"].ToString();
            String NIVEL_RECEP_RADIO_A = dt.Rows[0]["NIVEL_RECEP_RADIO_A"].ToString();
            String NIVEL_RECEP_RADIO_B = dt.Rows[0]["NIVEL_RECEP_RADIO_B"].ToString();
            String NIVEL_RECEP_NOM_A = dt.Rows[0]["NIVEL_RECEP_NOM_A"].ToString();
            String NIVEL_RECEP_NOM_B = dt.Rows[0]["NIVEL_RECEP_NOM_B"].ToString();
            String PING_PTP_RADIO_A = dt.Rows[0]["PING_PTP_RADIO_A"].ToString();
            String PING_PTP_RADIO_B = dt.Rows[0]["PING_PTP_RADIO_B"].ToString();

            byte[] CONF_GEN_ENL_EST_A = (byte[])dt.Rows[0]["CONF_GEN_ENL_EST_A"];
            MemoryStream mCONF_GEN_ENL_EST_A = new MemoryStream(CONF_GEN_ENL_EST_A);
            byte[] CONF_GEN_ENL_EST_B = (byte[])dt.Rows[0]["CONF_GEN_ENL_EST_B"];
            MemoryStream mCONF_GEN_ENL_EST_B = new MemoryStream(CONF_GEN_ENL_EST_B);
            byte[] CONF_LAN_EST_A01 = (byte[])dt.Rows[0]["CONF_LAN_EST_A01"];
            MemoryStream mCONF_LAN_EST_A01 = new MemoryStream(CONF_LAN_EST_A01);
            byte[] CONF_LAN_EST_A02 = (byte[])dt.Rows[0]["CONF_LAN_EST_A02"];
            MemoryStream mCONF_LAN_EST_A02 = new MemoryStream(CONF_LAN_EST_A02);
            byte[] CONF_LAN_EST_A03 = (byte[])dt.Rows[0]["CONF_LAN_EST_A03"];
            MemoryStream mCONF_LAN_EST_A03 = new MemoryStream(CONF_LAN_EST_A03);
            byte[] CONF_LAN_EST_B01 = (byte[])dt.Rows[0]["CONF_LAN_EST_B01"];
            MemoryStream mCONF_LAN_EST_B01 = new MemoryStream(CONF_LAN_EST_B01);
            byte[] CONF_LAN_EST_B02 = (byte[])dt.Rows[0]["CONF_LAN_EST_B02"];
            MemoryStream mCONF_LAN_EST_B02 = new MemoryStream(CONF_LAN_EST_B02);
            byte[] CONF_LAN_EST_B03 = (byte[])dt.Rows[0]["CONF_LAN_EST_B03"];
            MemoryStream mCONF_LAN_EST_B03 = new MemoryStream(CONF_LAN_EST_B03);
            byte[] CONF_ETHER_SWITCH_EST_A01 = (byte[])dt.Rows[0]["CONF_ETHER_SWITCH_EST_A01"];
            MemoryStream mCONF_ETHER_SWITCH_EST_A01 = new MemoryStream(CONF_ETHER_SWITCH_EST_A01);
            byte[] CONF_ETHER_SWITCH_EST_A02 = (byte[])dt.Rows[0]["CONF_ETHER_SWITCH_EST_A02"];
            MemoryStream mCONF_ETHER_SWITCH_EST_A02 = new MemoryStream(CONF_ETHER_SWITCH_EST_A02);
            byte[] CONF_ETHER_SWITCH_EST_B01 = (byte[])dt.Rows[0]["CONF_ETHER_SWITCH_EST_B01"];
            MemoryStream mCONF_ETHER_SWITCH_EST_B01 = new MemoryStream(CONF_ETHER_SWITCH_EST_B01);
            byte[] CONF_ETHER_SWITCH_EST_B02 = (byte[])dt.Rows[0]["CONF_ETHER_SWITCH_EST_B02"];
            MemoryStream mCONF_ETHER_SWITCH_EST_B02 = new MemoryStream(CONF_ETHER_SWITCH_EST_B02);

            byte[] CONF_IP_ESTAC_A = (byte[])dt.Rows[0]["CONF_IP_ESTAC_A"];
            MemoryStream mCONF_IP_ESTAC_A = new MemoryStream(CONF_IP_ESTAC_A);
            byte[] CONF_IP_ESTAC_B = (byte[])dt.Rows[0]["CONF_IP_ESTAC_B"];
            MemoryStream mCONF_IP_ESTAC_B = new MemoryStream(CONF_IP_ESTAC_B);
            #endregion

            #region Longitud SFTP

            //String VALOR_B_ESTAC_A = dt.Rows[0]["VALOR_B_ESTAC_A"].ToString();
            //String VALOR_C_ESTAC_A = dt.Rows[0]["VALOR_C_ESTAC_A"].ToString();
            //String VALOR_D_ESTAC_A = dt.Rows[0]["VALOR_D_ESTAC_A"].ToString();
            //String VALOR_E_ESTAC_A = dt.Rows[0]["VALOR_E_ESTAC_A"].ToString();
            //String VALOR_B_ESTAC_B = dt.Rows[0]["VALOR_B_ESTAC_B"].ToString();
            //String VALOR_C_ESTAC_B = dt.Rows[0]["VALOR_C_ESTAC_B"].ToString();
            //String VALOR_D_ESTAC_B = dt.Rows[0]["VALOR_D_ESTAC_B"].ToString();
            //String VALOR_E_ESTAC_B = dt.Rows[0]["VALOR_E_ESTAC_B"].ToString();

            #endregion

            #region Asignaciones y Observaciones

            String SWITCH_ROUTER_A = dt.Rows[0]["SWITCH_ROUTER_A"].ToString();
            String SWITCH_ROUTER_B = dt.Rows[0]["SWITCH_ROUTER_B"].ToString();
            String CAP_BREAKER_ASIG_EST_A = dt.Rows[0]["CAP_BREAKER_ASIG_EST_A"].ToString();
            String VOLT_DC_ESTAC_A = dt.Rows[0]["VOLT_DC_ESTAC_A"].ToString();
            String POS_BREAKER_ASIG_ESTAC_A = dt.Rows[0]["POS_BREAKER_ASIG_ESTAC_A"].ToString();
            String POS_BARRA_ATERRA_ESTA_A = dt.Rows[0]["POS_BARRA_ATERRA_ESTA_A"].ToString();
            String CAP_BREAKER_ASIG_EST_B = dt.Rows[0]["CAP_BREAKER_ASIG_EST_B"].ToString();
            String VOLT_DC_ESTAC_B = dt.Rows[0]["VOLT_DC_ESTAC_B"].ToString();
            String POS_BREAKER_ASIG_ESTAC_B = dt.Rows[0]["POS_BREAKER_ASIG_ESTAC_B"].ToString();
            String POS_BARRA_ATERRA_ESTA_B = dt.Rows[0]["POS_BARRA_ATERRA_ESTA_B"].ToString();

            #endregion

            #region Calculo Propagacion

            byte[] INGENIERIA = (byte[])dt.Rows[0]["INGENIERIA"];
            MemoryStream mINGENIERIA = new MemoryStream(INGENIERIA);
            byte[] PERFIL = (byte[])dt.Rows[0]["PERFIL"];
            MemoryStream mPERFIL = new MemoryStream(PERFIL);

            #endregion

            #region Pruebas de Interferencia

            byte[] PANT_RADIO_ESTAC_A = (byte[])dt.Rows[0]["PANT_RADIO_ESTAC_A"];
            MemoryStream mPANT_RADIO_ESTAC_A = new MemoryStream(PANT_RADIO_ESTAC_A);
            byte[] PANT_RADIO_ESTAC_B = (byte[])dt.Rows[0]["PANT_RADIO_ESTAC_B"];
            MemoryStream mPANT_RADIO_ESTAC_B = new MemoryStream(PANT_RADIO_ESTAC_B);

            #endregion

            #region Serie Equipos Fotos

            byte[] SERIE_ANT_ESTAC_A = (byte[])dt.Rows[0]["SERIE_ANT_ESTAC_A"];
            MemoryStream mSERIE_ANT_ESTAC_A = new MemoryStream(SERIE_ANT_ESTAC_A);
            byte[] SERIE_ODU_ESTAC_A = (byte[])dt.Rows[0]["SERIE_ODU_ESTAC_A"];
            MemoryStream mSERIE_ODU_ESTAC_A = new MemoryStream(SERIE_ODU_ESTAC_A);
            byte[] SERIE_POE_ESTAC_A = (byte[])dt.Rows[0]["SERIE_POE_ESTAC_A"];
            MemoryStream mSERIE_POE_ESTAC_A = new MemoryStream(SERIE_POE_ESTAC_A);

            CMM4BE CMM4A = new CMM4BE();
            List<CMM4BE> lstCMM4A = new List<CMM4BE>();
            CMM4A.Nodo.IdNodo = NODO_A;
            lstCMM4A = CMM4BL.ListarCMM4(CMM4A);


            CMM4BE CMM4B = new CMM4BE();
            List<CMM4BE> lstCMM4B = new List<CMM4BE>();
            CMM4A.Nodo.IdNodo = NODO_B;
            lstCMM4A = CMM4BL.ListarCMM4(CMM4B);

            byte[] SERIE_ANT_ESTAC_B = (byte[])dt.Rows[0]["SERIE_ANT_ESTAC_B"];
            MemoryStream mSERIE_ANT_ESTAC_B = new MemoryStream(SERIE_ANT_ESTAC_B);
            byte[] SERIE_ODU_ESTAC_B = (byte[])dt.Rows[0]["SERIE_ODU_ESTAC_B"];
            MemoryStream mSERIE_ODU_ESTAC_B = new MemoryStream(SERIE_ODU_ESTAC_B);
            byte[] SERIE_POE_ESTAC_B = (byte[])dt.Rows[0]["SERIE_POE_ESTAC_B"];
            MemoryStream mSERIE_POE_ESTAC_B = new MemoryStream(SERIE_POE_ESTAC_B);

            #endregion

            #region Reporte Fotografico

            #region Estacion A
            byte[] FOTO1_PAN_ESTAC_A = (byte[])dt.Rows[0]["FOTO1_PAN_ESTAC_A"];
            MemoryStream mFOTO1_PAN_ESTAC_A = new MemoryStream(FOTO1_PAN_ESTAC_A);
            byte[] FOTO2_POS_ANT_INST_TORRE_A = (byte[])dt.Rows[0]["FOTO2_POS_ANT_INST_TORRE_A"];
            MemoryStream mFOTO2_POS_ANT_INST_TORRE_A = new MemoryStream(FOTO2_POS_ANT_INST_TORRE_A);
            byte[] FOTO3_FOTO_ANTENA_ODU_ESTAC_A = (byte[])dt.Rows[0]["FOTO3_FOTO_ANTENA_ODU_ESTAC_A"];
            MemoryStream mFOTO3_FOTO_ANTENA_ODU_ESTAC_A = new MemoryStream(FOTO3_FOTO_ANTENA_ODU_ESTAC_A);
            byte[] FOTO5_ENGRAS_PERNOS_ESTAC_A = (byte[])dt.Rows[0]["FOTO5_ENGRAS_PERNOS_ESTAC_A"];
            MemoryStream mFOTO5_ENGRAS_PERNOS_ESTAC_A = new MemoryStream(FOTO5_ENGRAS_PERNOS_ESTAC_A);
            byte[] FOTO6_SILC_CONECT_ESTAC_A = (byte[])dt.Rows[0]["FOTO6_SILC_CONECT_ESTAC_A"];
            MemoryStream mFOTO6_SILC_CONECT_ESTAC_A = new MemoryStream(FOTO6_SILC_CONECT_ESTAC_A);
            byte[] FOTO7_1_ATERR_ODU_TORRE_ESTAC_A = (byte[])dt.Rows[0]["FOTO7_1_ATERR_ODU_TORRE_ESTAC_A"];
            MemoryStream mFOTO7_1_ATERR_ODU_TORRE_ESTAC_A = new MemoryStream(FOTO7_1_ATERR_ODU_TORRE_ESTAC_A);
            byte[] FOTO7_2_ATERR_ODU_TORRE_ESTAC_A = (byte[])dt.Rows[0]["FOTO7_2_ATERR_ODU_TORRE_ESTAC_A"];
            MemoryStream mFOTO7_2_ATERR_ODU_TORRE_ESTAC_A = new MemoryStream(FOTO7_2_ATERR_ODU_TORRE_ESTAC_A);
            byte[] FOTO8_RECORRI_SFTP_ESTAC_A = (byte[])dt.Rows[0]["FOTO8_RECORRI_SFTP_ESTAC_A"];
            MemoryStream mFOTO8_RECORRI_SFTP_ESTAC_A = new MemoryStream(FOTO8_RECORRI_SFTP_ESTAC_A);
            byte[] FOTO9_1__SFTP_OUT_1_ESTAC_A = (byte[])dt.Rows[0]["FOTO9_1__SFTP_OUT_1_ESTAC_A"];
            MemoryStream mFOTO9_1__SFTP_OUT_1_ESTAC_A = new MemoryStream(FOTO9_1__SFTP_OUT_1_ESTAC_A);
            byte[] FOTO9_2_SFTP_OUT_2_ESTAC_A = (byte[])dt.Rows[0]["FOTO9_2_SFTP_OUT_2_ESTAC_A"];
            MemoryStream mFOTO9_2_SFTP_OUT_2_ESTAC_A = new MemoryStream(FOTO9_2_SFTP_OUT_2_ESTAC_A);
            byte[] FOTO21__EITQU_POE_CMM4_ESTAC_A = (byte[])dt.Rows[0]["FOTO21__EITQU_POE_CMM4_ESTAC_A"];
            MemoryStream mFOTO21__EITQU_POE_CMM4_ESTAC_A = new MemoryStream(FOTO21__EITQU_POE_CMM4_ESTAC_A);
            byte[] FOTO22__PAN_RACK_ESTAC_A = (byte[])dt.Rows[0]["FOTO22__PAN_RACK_ESTAC_A"];
            MemoryStream mFOTO22__PAN_RACK_ESTAC_A = new MemoryStream(FOTO22__PAN_RACK_ESTAC_A);
            byte[] FOTO23_1_ATERRAM_POE_ESTAC_A = (byte[])dt.Rows[0]["FOTO23_1_ATERRAM_POE_ESTAC_A"];
            MemoryStream mFOTO23_1_ATERRAM_POE_ESTAC_A = new MemoryStream(FOTO23_1_ATERRAM_POE_ESTAC_A);
            byte[] FOTO23_2_ATERRAM_POE_ESTAC_A = (byte[])dt.Rows[0]["FOTO23_2_ATERRAM_POE_ESTAC_A"];
            MemoryStream mFOTO23_2_ATERRAM_POE_ESTAC_A = new MemoryStream(FOTO23_2_ATERRAM_POE_ESTAC_A);
            byte[] FOTO24_1_ENERG_POE_ETIQ_ESTAC_A = (byte[])dt.Rows[0]["FOTO24_1_ENERG_POE_ETIQ_ESTAC_A"];
            MemoryStream mFOTO24_1_ENERG_POE_ETIQ_ESTAC_A = new MemoryStream(FOTO24_1_ENERG_POE_ETIQ_ESTAC_A);
            byte[] FOTO24_2_ENERG_POE_ETIQ_ESTAC_A = (byte[])dt.Rows[0]["FOTO24_2_ENERG_POE_ETIQ_ESTAC_A"];
            MemoryStream mFOTO24_2_ENERG_POE_ETIQ_ESTAC_A = new MemoryStream(FOTO24_2_ENERG_POE_ETIQ_ESTAC_A);
            byte[] FOTO25_1_PATCH_POE_ESTAC_A = (byte[])dt.Rows[0]["FOTO25_1_PATCH_POE_ESTAC_A"];
            MemoryStream mFOTO25_1_PATCH_POE_ESTAC_A = new MemoryStream(FOTO25_1_PATCH_POE_ESTAC_A);
            byte[] FOTO25_2_PATCH_POE_ESTAC_A = (byte[])dt.Rows[0]["FOTO25_2_PATCH_POE_ESTAC_A"];
            MemoryStream mFOTO25_2_PATCH_POE_ESTAC_A = new MemoryStream(FOTO25_2_PATCH_POE_ESTAC_A);
            byte[] FOTO27_POE_CMM4_ESTAC_A = (byte[])dt.Rows[0]["FOTO27_POE_CMM4_ESTAC_A"];
            MemoryStream mFOTO27_POE_CMM4_ESTAC_A = new MemoryStream(FOTO27_POE_CMM4_ESTAC_A);
            #endregion

            #region Estacion B
            byte[] FOTO1_PAN_ESTAC_B = (byte[])dt.Rows[0]["FOTO1_PAN_ESTAC_B"];
            MemoryStream mFOTO1_PAN_ESTAC_B = new MemoryStream(FOTO1_PAN_ESTAC_B);
            byte[] FOTO2_POS_ANT_INST_TORRE_B = (byte[])dt.Rows[0]["FOTO2_POS_ANT_INST_TORRE_B"];
            MemoryStream mFOTO2_POS_ANT_INST_TORRE_B = new MemoryStream(FOTO2_POS_ANT_INST_TORRE_B);
            byte[] FOTO3_FOTO_ANTENA_ODU_ESTAC_B = (byte[])dt.Rows[0]["FOTO3_FOTO_ANTENA_ODU_ESTAC_B"];
            MemoryStream mFOTO3_FOTO_BNTENA_ODU_ESTAC_B = new MemoryStream(FOTO3_FOTO_ANTENA_ODU_ESTAC_B);
            byte[] FOTO5_ENGRAS_PERNOS_ESTAC_B = (byte[])dt.Rows[0]["FOTO5_ENGRAS_PERNOS_ESTAC_B"];
            MemoryStream mFOTO5_ENGRAS_PERNOS_ESTAC_B = new MemoryStream(FOTO5_ENGRAS_PERNOS_ESTAC_B);
            byte[] FOTO6_SILC_CONECT_ESTAC_B = (byte[])dt.Rows[0]["FOTO6_SILC_CONECT_ESTAC_B"];
            MemoryStream mFOTO6_SILC_CONECT_ESTAC_B = new MemoryStream(FOTO6_SILC_CONECT_ESTAC_B);
            byte[] FOTO7_1_ATERR_ODU_TORRE_ESTAC_B = (byte[])dt.Rows[0]["FOTO7_1_ATERR_ODU_TORRE_ESTAC_B"];
            MemoryStream mFOTO7_1_ATERR_ODU_TORRE_ESTAC_B = new MemoryStream(FOTO7_1_ATERR_ODU_TORRE_ESTAC_B);
            byte[] FOTO7_2_ATERR_ODU_TORRE_ESTAC_B = (byte[])dt.Rows[0]["FOTO7_2_ATERR_ODU_TORRE_ESTAC_B"];
            MemoryStream mFOTO7_2_ATERR_ODU_TORRE_ESTAC_B = new MemoryStream(FOTO7_2_ATERR_ODU_TORRE_ESTAC_B);
            byte[] FOTO8_RECORRI_SFTP_ESTAC_B = (byte[])dt.Rows[0]["FOTO8_RECORRI_SFTP_ESTAC_B"];
            MemoryStream mFOTO8_RECORRI_SFTP_ESTAC_B = new MemoryStream(FOTO8_RECORRI_SFTP_ESTAC_B);
            byte[] FOTO9_1_SFTP_OUT_1_ESTAC_B = (byte[])dt.Rows[0]["FOTO9_1_SFTP_OUT_1_ESTAC_B"];
            MemoryStream mFOTO9_1_SFTP_OUT_1_ESTAC_B = new MemoryStream(FOTO9_1_SFTP_OUT_1_ESTAC_B);
            byte[] FOTO9_2_SFTP_OUT_2_ESTAC_B = (byte[])dt.Rows[0]["FOTO9_2_SFTP_OUT_2_ESTAC_B"];
            MemoryStream mFOTO9_2_SFTP_OUT_2_ESTAC_B = new MemoryStream(FOTO9_2_SFTP_OUT_2_ESTAC_B);
            byte[] FOTO21__EITQU_POE_CMM4_ESTAC_B = (byte[])dt.Rows[0]["FOTO21__EITQU_POE_CMM4_ESTAC_B"];
            MemoryStream mFOTO21__EITQU_POE_CMM4_ESTAC_B = new MemoryStream(FOTO21__EITQU_POE_CMM4_ESTAC_B);
            byte[] FOTO22_PAN_RACK_ESTAC_B = (byte[])dt.Rows[0]["FOTO22_PAN_RACK_ESTAC_B"];
            MemoryStream mFOTO22_PAN_RACK_ESTAC_B = new MemoryStream(FOTO22_PAN_RACK_ESTAC_B);
            byte[] FOTO23_1_ATERRAM_POE_ESTAC_B = (byte[])dt.Rows[0]["FOTO23_1_ATERRAM_POE_ESTAC_B"];
            MemoryStream mFOTO23_1_ATERRAM_POE_ESTAC_B = new MemoryStream(FOTO23_1_ATERRAM_POE_ESTAC_B);
            byte[] FOTO23_2_ATERRAM_POE_ESTAC_B = (byte[])dt.Rows[0]["FOTO23_2_ATERRAM_POE_ESTAC_B"];
            MemoryStream mFOTO23_2_ATERRAM_POE_ESTAC_B = new MemoryStream(FOTO23_2_ATERRAM_POE_ESTAC_B);
            byte[] FOTO24_1_ENERG_POE_ETIQ_ESTAC_B = (byte[])dt.Rows[0]["FOTO24_1_ENERG_POE_ETIQ_ESTAC_B"];
            MemoryStream mFOTO24_1_ENERG_POE_ETIQ_ESTAC_B = new MemoryStream(FOTO24_1_ENERG_POE_ETIQ_ESTAC_B);
            byte[] FOTO24_2_ENERG_POE_ETIQ_ESTAC_B = (byte[])dt.Rows[0]["FOTO24_2_ENERG_POE_ETIQ_ESTAC_B"];
            MemoryStream mFOTO24_2_ENERG_POE_ETIQ_ESTAC_B = new MemoryStream(FOTO24_2_ENERG_POE_ETIQ_ESTAC_B);
            byte[] FOTO25_1_PATCH_POE_ESTAC_B = (byte[])dt.Rows[0]["FOTO25_1_PATCH_POE_ESTAC_B"];
            MemoryStream mFOTO25_1_PATCH_POE_ESTAC_B = new MemoryStream(FOTO25_1_PATCH_POE_ESTAC_B);
            byte[] FOTO25_2_PATCH_POE_ESTAC_B = (byte[])dt.Rows[0]["FOTO25_2_PATCH_POE_ESTAC_B"];
            MemoryStream mFOTO25_2_PATCH_POE_ESTAC_B = new MemoryStream(FOTO25_2_PATCH_POE_ESTAC_B);
            byte[] FOTO27_POE_CMM4_ESTAC_B = (byte[])dt.Rows[0]["FOTO27_POE_CMM4_ESTAC_B"];
            MemoryStream mFOTO27_POE_CMM4_ESTAC_B = new MemoryStream(FOTO27_POE_CMM4_ESTAC_B);
            #endregion

            #endregion

            #region Datos Generales Nodo A

            String UBIGEO_A = dt.Rows[0]["UBIGEO_A"].ToString();
            String SERIE_PTP_450_A = dt.Rows[0]["SERIE_PTP_450_A"].ToString();
            String FRECUENCIA_A = dt.Rows[0]["FRECUENCIA_A"].ToString();
            String PIRE_EIRP_A = dt.Rows[0]["PIRE_EIRP_A"].ToString();
            String ANTENA_MARCA_MODELO_A = dt.Rows[0]["ANTENA_MARCA_MODELO_A"].ToString();
            String GANANCIA_ANTENA_A = dt.Rows[0]["GANANCIA_ANTENA_A"].ToString();
            //String ALTURA_ANTENA_A = ds.Tables[0].Rows[0]["ALTURA_ANTENA_A"].ToString();
            String ELEVACION_A = dt.Rows[0]["ELEVACION_A"].ToString();
            String ALTITUD_msnm_A = dt.Rows[0]["ALTITUD_msnm_A"].ToString();



            #endregion

            #region Datos generales Nodo B

            String UBIGEO_B = dt.Rows[0]["UBIGEO_B"].ToString();
            String SERIE_PTP_450_B = dt.Rows[0]["SERIE_PTP_450_B"].ToString();
            String FRECUENCIA_B = dt.Rows[0]["FRECUENCIA_B"].ToString();
            String PIRE_EIRP_B = dt.Rows[0]["PIRE_EIRP_B"].ToString();
            String ANTENA_MARCA_MODELO_B = dt.Rows[0]["ANTENA_MARCA_MODELO_B"].ToString();
            String GANANCIA_ANTENA_B = dt.Rows[0]["GANANCIA_ANTENA_B"].ToString();
            //  String ALTURA_ANTENA_B = ds.Tables[0].Rows[0]["ALTURA_ANTENA_B"].ToString();
            String ELEVACION_B = dt.Rows[0]["ELEVACION_B"].ToString();
            String ALTITUD_msnm_B = dt.Rows[0]["ALTITUD_msnm_B"].ToString();

            #endregion

            #endregion

            // String usuarioWindows = Environment.UserName;
            String excelGenerado = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";

            File.Copy(rutaPlantilla, excelGenerado, true);

            #region Ingresando Valores

            #region Caratula

            ExcelToolsBL.UpdateCell(excelGenerado, "Carátula", NOMBRE_NODO_A, 15, "C");
            ExcelToolsBL.UpdateCell(excelGenerado, "Carátula", NODO_A, 18, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "Carátula", NODO_B, 18, "F");
            ExcelToolsBL.UpdateCell(excelGenerado, "Carátula", NOMBRE_NODO_A, 19, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "Carátula", NOMBRE_NODO_A, 19, "F");
            ExcelToolsBL.UpdateCell(excelGenerado, "Carátula", TIPO_NODO_A, 20, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "Carátula", TIPO_NODO_B, 20, "F");
            ExcelToolsBL.UpdateCell(excelGenerado, "Carátula", FRECUENCIA_A, 22, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "Carátula", FECHA, 24, "E");

            #endregion

            #region Configuracion y Mediciones

            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", TIPO_NODO_A, 16, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", TIPO_NODO_B, 16, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", NODO_A, 17, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", NODO_B, 17, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", NOMBRE_NODO_A, 18, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", NOMBRE_NODO_B, 18, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", DIREC_ESTACION_A, 19, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", DIREC_ESTACION_B, 19, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", DISTRITO_A, 20, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", DISTRITO_B, 20, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", PROVINCIA_A, 21, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", PROVINCIA_B, 21, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", DEPARTAMENTO_A, 22, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", DEPARTAMENTO_B, 22, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", LATITUD_A, 23, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", LATITUD_B, 23, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", LONGITUD_A, 24, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", LONGITUD_B, 24, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", ALTURAmsnm_A, 25, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", ALTURAmsnm_B, 25, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", REF_UBIC_EST_A, 26, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", REF_UBIC_EST_B, 26, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", ALTURA_TORRE_A, 28, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", ALTURA_TORRE_B, 28, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", DISTANCIA_A_B, 29, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", MODULACION, 31, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", MODULACION, 31, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", FRECUENCIA_TX_A, 32, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", FRECUENCIA_TX_B, 32, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", FRECUENCIA_TX_B, 33, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", FRECUENCIA_TX_A, 33, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", CANAL_ESTAC_A, 34, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", CANAL_ESTAC_B, 34, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", VELOCIDAD_HABILITADA, 37, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", VELOCIDAD_HABILITADA, 37, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", ANCHO_BANDA, 38, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", MODEL_ANTENA_A, 39, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", MODEL_ANTENA_B, 39, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", DIAMETRO_A, 40, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", DIAMETRO_B, 40, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", ALTURA_ANTENA_A, 41, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", ALTURA_ANTENA_B, 41, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", POLARIZACION_A, 42, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", POLARIZACION_B, 42, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", AZIMUT_A, 43, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", AZIMUT_B, 43, "j");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", "Polarizacion: " + POLARIZACION_A, 50, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", "Modulación:" + MODULACION + "QAM", 50, "K");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", IP_NODO_A_1, 55, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", IP_NODO_A_2, 55, "F");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", IP_NODO_A_3, 55, "G");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", IP_NODO_A_4, 55, "H");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", IP_NODO_B_1, 55, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", IP_NODO_B_2, 55, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", IP_NODO_B_3, 55, "K");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", IP_NODO_B_4, 55, "L");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", IP_NODO_A_1, 56, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", IP_NODO_A_2, 56, "F");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", IP_NODO_A_3, 56, "G");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", IP_NODO_A_4, 56, "H");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", IP_NODO_B_1, 56, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", IP_NODO_B_2, 56, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", IP_NODO_B_3, 56, "K");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", IP_NODO_B_4, 56, "L");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", DEFAULT_GATE_AB_1, 58, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", DEFAULT_GATE_AB_2, 58, "F");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", DEFAULT_GATE_AB_3, 58, "G");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", DEFAULT_GATE_AB_4, 58, "H");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", DEFAULT_GATE_AB_1, 58, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", DEFAULT_GATE_AB_2, 58, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", DEFAULT_GATE_AB_3, 58, "K");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", DEFAULT_GATE_AB_4, 58, "L");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", POTENCIA_A, 71, "K");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", POTENCIA_B, 71, "M");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", POTENCIA_A, 72, "K");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", POTENCIA_B, 72, "M");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", MARGEN_DES_A, 73, "K");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", MARGEN_DES_B, 73, "M");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", NIVEL_UMBRAL_A_B, 74, "K");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", NIVEL_UMBRAL_A_B, 74, "M");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", NIVEL_RECEP_NOM_A, 77, "K");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", NIVEL_RECEP_NOM_B, 77, "M");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", NIVEL_RECEP_RADIO_A, 78, "K");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", NIVEL_RECEP_RADIO_B, 78, "M");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", PING_PTP_RADIO_A + " ms", 87, "K");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", PING_PTP_RADIO_B + " ms", 87, "M");

            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Mediciones", mCONF_GEN_ENL_EST_A, "", 94,3,724,300);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Mediciones", mCONF_GEN_ENL_EST_B, "", 109,3,725, 339);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Mediciones", mCONF_LAN_EST_A01, "", 128,3,257,278);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Mediciones", mCONF_LAN_EST_A02, "", 128,5,247,277);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Mediciones", mCONF_LAN_EST_A03, "", 128,11,217, 278);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Mediciones", mCONF_LAN_EST_B01, "", 144, 3,256,238);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Mediciones", mCONF_LAN_EST_B02, "", 144, 5,254,240);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Mediciones", mCONF_LAN_EST_B03, "", 144,11, 212, 237);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Mediciones", mCONF_ETHER_SWITCH_EST_A01, "", 163, 3, 370, 301);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Mediciones", mCONF_ETHER_SWITCH_EST_A02, "",163,8, 352, 306);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Mediciones", mCONF_ETHER_SWITCH_EST_B01, "", 179, 3, 373, 301);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Mediciones", mCONF_ETHER_SWITCH_EST_B02, "", 179, 8, 352, 302);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Mediciones", mCONF_IP_ESTAC_A, "", 198, 3,724, 239);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Mediciones", mCONF_IP_ESTAC_B, "", 214, 3,723, 341);

            #endregion

            #region Materiales A

            ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales A", NOMBRE_NODO_A + " - " + NOMBRE_NODO_B, 12, "C");
            ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales A", NOMBRE_NODO_A, 12, "F");

            foreach (DataRow dr in dt1.Rows)
            {
                String EQUIPO = dr["EQUIPO"].ToString();
                String MARCA = dr["MARCA"].ToString();
                String MODELO = dr["MODELO"].ToString();
                String Nro_SERIE = dr["Nro_SERIE"].ToString();

                int ind = Convert.ToInt32(dt1.Rows.IndexOf(dr));

                ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales A", Convert.ToString(ind + 1), 17 + ind, "B");
                ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales A", EQUIPO, 17 + ind, "C");
                ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales A", "1", 17 + ind, "D");
                ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales A", MARCA, 17 + ind, "E");
                ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales A", MODELO, 17 + ind, "F");
                ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales A", Nro_SERIE, 17 + ind, "G");

            }

            foreach (DataRow dr in dt2.Rows)
            {
                String DESCRIPCION = dr["DESCRIPCION"].ToString();
                String UNIDAD = dr["UNIDAD"].ToString();
                String CANTIDAD = dr["CANTIDAD"].ToString();

                int ind = Convert.ToInt32(dt2.Rows.IndexOf(dr));

                ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales A", Convert.ToString(ind + 1), 32 + ind, "B");
                ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales A", DESCRIPCION, 32 + ind, "C");
                ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales A", UNIDAD, 32 + ind, "E");
                ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales A", CANTIDAD, 32 + ind, "F");

            }

            #endregion

            #region Materiales B

            ExcelToolsBL.UpdateCell(excelGenerado, "3 Materiales B", NOMBRE_NODO_A + " - " + NOMBRE_NODO_B, 12, "C");
            ExcelToolsBL.UpdateCell(excelGenerado, "3 Materiales B", NOMBRE_NODO_B, 12, "F");

            foreach (DataRow dr in dt3.Rows)
            {
                String EQUIPO = dr["EQUIPO"].ToString();
                String MARCA = dr["MARCA"].ToString();
                String MODELO = dr["MODELO"].ToString();
                String Nro_SERIE = dr["Nro_SERIE"].ToString();

                int ind = Convert.ToInt32(dt3.Rows.IndexOf(dr));

                ExcelToolsBL.UpdateCell(excelGenerado, "3 Materiales B", Convert.ToString(ind + 1), 17 + ind, "B");
                ExcelToolsBL.UpdateCell(excelGenerado, "3 Materiales B", EQUIPO, 17 + ind, "C");
                ExcelToolsBL.UpdateCell(excelGenerado, "3 Materiales B", "1", 17 + ind, "D");
                ExcelToolsBL.UpdateCell(excelGenerado, "3 Materiales B", MARCA, 17 + ind, "E");
                ExcelToolsBL.UpdateCell(excelGenerado, "3 Materiales B", MODELO, 17 + ind, "F");
                ExcelToolsBL.UpdateCell(excelGenerado, "3 Materiales B", Nro_SERIE, 17 + ind, "G");

            }

            foreach (DataRow dr in dt4.Rows)
            {
                String DESCRIPCION = dr["DESCRIPCION"].ToString();
                String UNIDAD = dr["UNIDAD"].ToString();
                String CANTIDAD = dr["CANTIDAD"].ToString();

                int ind = Convert.ToInt32(dt4.Rows.IndexOf(dr));

                ExcelToolsBL.UpdateCell(excelGenerado, "3 Materiales B", Convert.ToString(ind + 1), 32 + ind, "B");
                ExcelToolsBL.UpdateCell(excelGenerado, "3 Materiales B", DESCRIPCION, 32 + ind, "C");
                ExcelToolsBL.UpdateCell(excelGenerado, "3 Materiales B", UNIDAD, 32 + ind, "E");
                ExcelToolsBL.UpdateCell(excelGenerado, "3 Materiales B", CANTIDAD, 32 + ind, "F");

            }

            #endregion

            #region Longitud SFTP

            ExcelToolsBL.UpdateCell(excelGenerado, "4 Longitud SFTP", NOMBRE_NODO_A + " - " + NOMBRE_NODO_B, 13, "D");

            ExcelToolsBL.UpdateCell(excelGenerado, "4 Longitud SFTP", DIREC_ESTACION_A, 20, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "4 Longitud SFTP", ALTURA_TORRE_A, 20, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "4 Longitud SFTP", ALTURA_ANTENA_A, 20, "F");
            ExcelToolsBL.UpdateCell(excelGenerado, "4 Longitud SFTP", "3", 20, "G");
            ExcelToolsBL.UpdateCell(excelGenerado, "4 Longitud SFTP", "2,6", 20, "h");
            ExcelToolsBL.UpdateCell(excelGenerado, "4 Longitud SFTP", "8", 20, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "4 Longitud SFTP", "7", 20, "J");

            ExcelToolsBL.UpdateCell(excelGenerado, "4 Longitud SFTP", DIREC_ESTACION_B, 22, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "4 Longitud SFTP", ALTURA_TORRE_B, 22, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "4 Longitud SFTP", ALTURA_ANTENA_B, 22, "F");
            ExcelToolsBL.UpdateCell(excelGenerado, "4 Longitud SFTP", "3", 22, "G");
            ExcelToolsBL.UpdateCell(excelGenerado, "4 Longitud SFTP", "2,6", 22, "h");
            ExcelToolsBL.UpdateCell(excelGenerado, "4 Longitud SFTP", "8", 22, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "4 Longitud SFTP", "7", 22, "J");


            #endregion

            #region Asignaciones y Observaciones

            if (SWITCH_ROUTER_A.First().Equals("7"))
            {
                ExcelToolsBL.UpdateCell(excelGenerado, "5 Asignaciones y Observaciones", "Router Nokia Puerto" + " " + SWITCH_ROUTER_A.Last(), 17, "G");
            }
            else
            {
                ExcelToolsBL.UpdateCell(excelGenerado, "5 Asignaciones y Observaciones", "Omni Switch Puerto" + " " + SWITCH_ROUTER_A.Last(), 17, "G");
            }
            if (SWITCH_ROUTER_B.First().Equals("7"))
            {
                ExcelToolsBL.UpdateCell(excelGenerado, "5 Asignaciones y Observaciones", "Router Nokia Puerto" + " " + SWITCH_ROUTER_B.Last(), 22, "G");
            }
            else
            {
                ExcelToolsBL.UpdateCell(excelGenerado, "5 Asignaciones y Observaciones", "Omni Switch Puerto" + " " + SWITCH_ROUTER_B.Last(), 22, "G");
            }

            ExcelToolsBL.UpdateCell(excelGenerado, "5 Asignaciones y Observaciones", CAP_BREAKER_ASIG_EST_A, 27, "K");
            ExcelToolsBL.UpdateCell(excelGenerado, "5 Asignaciones y Observaciones", VOLT_DC_ESTAC_A, 28, "K");
            ExcelToolsBL.UpdateCell(excelGenerado, "5 Asignaciones y Observaciones", POS_BREAKER_ASIG_ESTAC_A, 29, "K");
            ExcelToolsBL.UpdateCell(excelGenerado, "5 Asignaciones y Observaciones", POS_BARRA_ATERRA_ESTA_A, 30, "K");
            ExcelToolsBL.UpdateCell(excelGenerado, "5 Asignaciones y Observaciones", CAP_BREAKER_ASIG_EST_B, 34, "K");
            ExcelToolsBL.UpdateCell(excelGenerado, "5 Asignaciones y Observaciones", VOLT_DC_ESTAC_B, 35, "K");
            ExcelToolsBL.UpdateCell(excelGenerado, "5 Asignaciones y Observaciones", POS_BREAKER_ASIG_ESTAC_B, 36, "K");
            ExcelToolsBL.UpdateCell(excelGenerado, "5 Asignaciones y Observaciones", POS_BARRA_ATERRA_ESTA_B, 37, "K");


            #endregion

            #region Calculo Propagacion

            ExcelToolsBL.AddImageDocument(false, excelGenerado, "6 Cálculo Propagacion", mINGENIERIA, "", 9, 6, 533, 498);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "6 Cálculo Propagacion", mPERFIL, "", 33, 5, 593, 373);

            #endregion

            #region Pruebas de Interferencia

            ExcelToolsBL.UpdateCell(excelGenerado, "8 Pruebas de Interferencia", NODO_A, 15, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "8 Pruebas de Interferencia", NODO_B, 15, "F");
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "8 Pruebas de Interferencia", mPANT_RADIO_ESTAC_A, "", 20, 3,723,341);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "8 Pruebas de Interferencia", mPANT_RADIO_ESTAC_B, "", 38, 3, 698,324);

            #endregion

            #region Serie Equipos Fotos

            ExcelToolsBL.AddImageDocument(false, excelGenerado, "9 Serie Equipos (fotos)", mSERIE_ANT_ESTAC_A, "", 14, 3, 120, 98);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "9 Serie Equipos (fotos)", mSERIE_ODU_ESTAC_A, "", 21, 3, 81, 113);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "9 Serie Equipos (fotos)", mSERIE_POE_ESTAC_A, "", 29, 3, 152, 117);


            if (!lstCMM4A.Count.Equals(0))
            {
                byte[] SERIE_CMM4_ESTAC_A = (byte[])dt.Rows[0]["SERIE_CMM4_ESTAC_A"];
                MemoryStream mSERIE_CMM4_ESTAC_A = new MemoryStream(SERIE_CMM4_ESTAC_A);
                byte[] SERIE_UGPS_ESTAC_A = (byte[])dt.Rows[0]["SERIE_UGPS_ESTAC_A"];
                MemoryStream mSERIE_UGPS_ESTAC_A = new MemoryStream(SERIE_UGPS_ESTAC_A);
                byte[] SERIE_CONVERSOR_ESTAC_A = (byte[])dt.Rows[0]["SERIE_CONVERSOR_ESTAC_A"];
                MemoryStream mSERIE_CONVERSOR_ESTAC_A = new MemoryStream(SERIE_CONVERSOR_ESTAC_A);

                ExcelToolsBL.AddImageDocument(false, excelGenerado, "9 Serie Equipos (fotos)", mSERIE_CMM4_ESTAC_A, "", 37, 3, 152, 117);
                ExcelToolsBL.AddImageDocument(false, excelGenerado, "9 Serie Equipos (fotos)", mSERIE_UGPS_ESTAC_A, "", 45, 3, 152, 117);
                ExcelToolsBL.AddImageDocument(false, excelGenerado, "9 Serie Equipos (fotos)", mSERIE_CONVERSOR_ESTAC_A, "", 52, 3, 152, 117);

            }

            if (!lstCMM4B.Count.Equals(0))
            {
                byte[] SERIE_CMM4_ESTAC_B = (byte[])dt.Rows[0]["SERIE_CMM4_ESTAC_B"];
                MemoryStream mSERIE_CMM4_ESTAC_B = new MemoryStream(SERIE_CMM4_ESTAC_B);
                byte[] SERIE_UGPS_ESTAC_B = (byte[])dt.Rows[0]["SERIE_UGPS_ESTAC_B"];
                MemoryStream mSERIE_UGPS_ESTAC_B = new MemoryStream(SERIE_UGPS_ESTAC_B);
                byte[] SERIE_CONVERSOR_ESTAC_B = (byte[])dt.Rows[0]["SERIE_CONVERSOR_ESTAC_B"];
                MemoryStream mSERIE_CONVERSOR_ESTAC_B = new MemoryStream(SERIE_CONVERSOR_ESTAC_B);

                ExcelToolsBL.AddImageDocument(false, excelGenerado, "9 Serie Equipos (fotos)", mSERIE_CMM4_ESTAC_B, "", 82, 3, 152, 117);
                ExcelToolsBL.AddImageDocument(false, excelGenerado, "9 Serie Equipos (fotos)", mSERIE_UGPS_ESTAC_B, "", 90, 3, 152, 117);
                ExcelToolsBL.AddImageDocument(false, excelGenerado, "9 Serie Equipos (fotos)", mSERIE_CONVERSOR_ESTAC_B, "", 97, 3, 152, 117);

            }


            ExcelToolsBL.AddImageDocument(false, excelGenerado, "9 Serie Equipos (fotos)", mSERIE_ANT_ESTAC_B, "", 59, 3, 179, 23);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "9 Serie Equipos (fotos)", mSERIE_ODU_ESTAC_B, "", 66, 3, 81, 113);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "9 Serie Equipos (fotos)", mSERIE_POE_ESTAC_B, "", 75, 3, 152, 117);


            #endregion

            #region Reporte Fotografico

            #region Estacion A

            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO1_PAN_ESTAC_A, "", 12, 3, 435, 593);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO2_POS_ANT_INST_TORRE_A, "", 12, 14, 435, 590);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO3_FOTO_ANTENA_ODU_ESTAC_A, "", 29, 3, 435, 348);
            // ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO5_ENGRAS_PERNOS_ESTAC_A, "", 29, 16, 189, 250);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO5_ENGRAS_PERNOS_ESTAC_A, "", 46, 3, 435, 333);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO6_SILC_CONECT_ESTAC_A, "", 46, 14, 433, 337);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO7_1_ATERR_ODU_TORRE_ESTAC_A, "", 62, 3, 435, 184);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO7_2_ATERR_ODU_TORRE_ESTAC_A, "",73,3,435, 343);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO8_RECORRI_SFTP_ESTAC_A, "", 62,14, 432, 529);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO9_1__SFTP_OUT_1_ESTAC_A, "",79,3,434, 222);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO9_2_SFTP_OUT_2_ESTAC_A,"",90,3,435,274);


            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO21__EITQU_POE_CMM4_ESTAC_A, "", 179, 3, 431, 343);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO22__PAN_RACK_ESTAC_A, "", 179, 14, 436, 344);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO23_1_ATERRAM_POE_ESTAC_A, "", 196, 3, 436, 263);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO23_2_ATERRAM_POE_ESTAC_A, "", 207, 3, 436, 290);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO24_1_ENERG_POE_ETIQ_ESTAC_A, "", 196, 14, 432, 265);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO24_2_ENERG_POE_ETIQ_ESTAC_A, "", 207, 14, 431, 292);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO25_1_PATCH_POE_ESTAC_A,"",213,3,435,259);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO25_2_PATCH_POE_ESTAC_A,"",224,3,434,311);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO27_POE_CMM4_ESTAC_A, "", 229, 3, 435, 330);

            #endregion

            #region Estacion B

            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO1_PAN_ESTAC_B,"",265,3,435,602);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO2_POS_ANT_INST_TORRE_B, "", 265, 14, 435,602);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO3_FOTO_BNTENA_ODU_ESTAC_B, "", 282, 3, 435, 348);
            // ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO5_ENGRAS_PERNOS_ESTAC_B, "", 29, 16, 189, 250);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO5_ENGRAS_PERNOS_ESTAC_B, "", 299, 3, 435, 333);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO6_SILC_CONECT_ESTAC_B, "", 299, 14, 433, 337);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO7_1_ATERR_ODU_TORRE_ESTAC_B,"",315,3,435, 188);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO7_2_ATERR_ODU_TORRE_ESTAC_B,"",326,3,435, 331);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO8_RECORRI_SFTP_ESTAC_B,"",315,14,434,521);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO9_1_SFTP_OUT_1_ESTAC_B,"",332,3,436,188);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO9_2_SFTP_OUT_2_ESTAC_B,"",343,3,435,162);


            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO21__EITQU_POE_CMM4_ESTAC_B,"",432,3,435,349);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO22_PAN_RACK_ESTAC_B,"",432,14,433,350);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO23_1_ATERRAM_POE_ESTAC_B,"",449,3,434,248);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO23_2_ATERRAM_POE_ESTAC_B,"",460,3,431,262);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO24_1_ENERG_POE_ETIQ_ESTAC_B, "", 449, 14, 432, 265);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO24_2_ENERG_POE_ETIQ_ESTAC_B, "", 460, 14, 431, 292);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO25_1_PATCH_POE_ESTAC_B,"",466,3,435,266);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO25_2_PATCH_POE_ESTAC_B,"",477,3,434,245);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO27_POE_CMM4_ESTAC_B,"",482,3,435,349);

            #endregion


            #endregion

            #region Datos Generales Nodo A

            ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", TIPO_NODO_A, 9, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", NODO_A, 12, "B");
            ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", NOMBRE_NODO_A, 12, "G");
            ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", NOMBRE_NODO_A, 17, "B");
            ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", UBIGEO_A, 17, "G");
            ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", DEPARTAMENTO_A, 20, "B");
            ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", PROVINCIA_A, 20, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", DISTRITO_A, 20, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", SERIE_PTP_450_A, 29, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", FRECUENCIA_A + " Mhz", 41, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", PIRE_EIRP_A + " dbm", 44, "I");

            if (ANTENA_MARCA_MODELO_A.Equals("INTEGRADA"))
            {
                ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", "Cambium Networks/5092HH", 46, "I");

            }
            else if (ANTENA_MARCA_MODELO_A.Equals("0.6"))
            {

                ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", "SHENGLU/SLU0652DD6B", 46, "I");
            }

            ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", GANANCIA_ANTENA_A, 47, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", ALTURA_ANTENA_A, 49, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", ELEVACION_A, 52, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", ALTITUD_msnm_A, 53, "I");

            foreach (DataRow dr in dt5.Rows)
            {
                String NODO_LOCAL = dr["NODO_LOCAL"].ToString();
                String NODO_REMOTO = dr["NODO_REMOTO"].ToString();
                String RSS_LOCAL = dr["RSS_LOCAL"].ToString();
                String RSS_REMOTO = dr["RSS_REMOTO"].ToString();
                String TIEMPO_PROM = dr["TIEMPO_PROM"].ToString();
                String CAP_SUBIDA = dr["CAP_SUBIDA"].ToString();
                String CAP_BAJADA = dr["CAP_BAJADA"].ToString();
                String DISTANCIA_metros = dr["DISTANCIA_metros"].ToString();

                ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", NODO_LOCAL, 60, "B");
                ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", NODO_REMOTO, 60, "C");
                ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", RSS_LOCAL, 60, "D");
                ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", RSS_REMOTO, 60, "E");
                ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", "20", 60, "F");
                ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", "16 QAM", 60, "G");
                ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", TIEMPO_PROM, 60, "H");
                ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", "UL " + CAP_SUBIDA + "/" + "DL " + CAP_BAJADA, 60, "I");
                ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", DISTANCIA_metros, 60, "K");
                ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", FRECUENCIA_A + "MHz", 60, "L");

            }


            #endregion

            #region Datos Generales Nodo B

            ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", TIPO_NODO_B, 9, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", NODO_B, 12, "B");
            ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", NOMBRE_NODO_B, 12, "G");
            ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", NOMBRE_NODO_B, 17, "B");
            ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", UBIGEO_B, 17, "G");
            ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", DEPARTAMENTO_B, 20, "B");
            ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", PROVINCIA_B, 20, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", DISTRITO_B, 20, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", SERIE_PTP_450_B, 29, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", FRECUENCIA_B + " Mhz", 41, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", PIRE_EIRP_B + " dbm", 44, "I");

            if (ANTENA_MARCA_MODELO_B.Equals("INTEGRADA"))
            {
                ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", "Cambium Networks/5092HH", 46, "I");

            }
            else if (ANTENA_MARCA_MODELO_B.Equals("0.6"))
            {

                ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", "SHENGLU/SLU0652DD6B", 46, "I");
            }

            ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", GANANCIA_ANTENA_B, 47, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", ALTURA_ANTENA_B, 49, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", ELEVACION_B, 52, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", ALTITUD_msnm_B, 53, "I");

            foreach (DataRow dr in dt6.Rows)
            {
                String NODO_LOCAL = dr["NODO_LOCAL"].ToString();
                String NODO_REMOTO = dr["NODO_REMOTO"].ToString();
                String RSS_LOCAL = dr["RSS_LOCAL"].ToString();
                String RSS_REMOTO = dr["RSS_REMOTO"].ToString();
                String TIEMPO_PROM = dr["TIEMPO_PROM"].ToString();
                String CAP_SUBIDA = dr["CAP_SUBIDA"].ToString();
                String CAP_BAJADA = dr["CAP_BAJADA"].ToString();
                String DISTANCIA_metros = dr["DISTANCIA_metros"].ToString();

                ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", NODO_LOCAL, 60, "B");
                ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", NODO_REMOTO, 60, "C");
                ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", RSS_LOCAL, 60, "D");
                ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", RSS_REMOTO, 60, "E");
                ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", "20", 60, "F");
                ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", "16 QAM", 60, "G");
                ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", TIEMPO_PROM, 60, "H");
                ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", "UL " + CAP_SUBIDA + "/" + "DL " + CAP_BAJADA, 60, "I");
                ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", DISTANCIA_metros, 60, "K");
                ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", FRECUENCIA_B + "MHz", 60, "L");

            }


            #endregion


            #endregion

            #region Ruta del Zip
            String rutaNodo = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo;
            if (Directory.Exists(rutaNodo))
            {
                String rutaParcial = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\ACCESO\\PTP_HC-0XX2 (alfo)\\";
                if (!Directory.Exists(rutaParcial)) Directory.CreateDirectory(rutaParcial);
                String rutaAlterna = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\ACCESO\\PTP_HC-0XX2 (alfo)\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";
                File.Copy(excelGenerado, rutaAlterna, true);
            }
            #endregion
        }

        public void ActaInstalacionPTPNoLicenciado(String IdNodo, String IdTarea, String valorCadena1, String rutaPlantilla)
        {
            DBBaseDatos baseDatosDA = new DBBaseDatos();

            baseDatosDA.Configurar();
            baseDatosDA.Conectar();
            DataTable dt = new DataTable();
            DataTable dt1 = new DataTable();
            DataTable dt2 = new DataTable();
            DataTable dt3 = new DataTable();
            DataTable dt4 = new DataTable();
            DataTable dt5 = new DataTable();
            DataTable dt6 = new DataTable();

            try
            {
                baseDatosDA.CrearComando("USP_R_ACTA_INSTALACION_PTP_NO_LIC", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                dt = baseDatosDA.EjecutarConsultaDataTable();

                baseDatosDA.CrearComando("USP_R_EQUIPAMIENTOS_PTP_NO_LIC_A", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA",IdTarea,true);
                dt1 = baseDatosDA.EjecutarConsultaDataTable();

                baseDatosDA.CrearComando("USP_R_MATERIALES_PTP_NO_LIC_A", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                dt2 = baseDatosDA.EjecutarConsultaDataTable();

                baseDatosDA.CrearComando("USP_R_EQUIPAMIENTOS_PTP_NO_LIC_B", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                dt3 = baseDatosDA.EjecutarConsultaDataTable();

                baseDatosDA.CrearComando("USP_R_MATERIALES_PTP_NO_LIC_B", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                dt4 = baseDatosDA.EjecutarConsultaDataTable();

                baseDatosDA.CrearComando("USP_R_MEDICION_ENLACE_PTP_NO_LIC_A", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                dt5 = baseDatosDA.EjecutarConsultaDataTable();

                baseDatosDA.CrearComando("USP_R_MEDICION_ENLACE_PTP_NO_LIC_B", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                dt6 = baseDatosDA.EjecutarConsultaDataTable();

            }
            catch (Exception)
            {

                throw;
            }
            finally
            {

                baseDatosDA.Desconectar();
                baseDatosDA = null;
            }


            #region Valores 

            #region Caratula

            //NOMBRE LOCALIDAD QUE LOCALIDAD SE PONDRIA LA A O LA B
            String NODO_A = dt.Rows[0]["NODO_A"].ToString();
            String NODO_B = dt.Rows[0]["NODO_B"].ToString();
            String NOMBRE_NODO_A = dt.Rows[0]["NOMBRE_NODO_A"].ToString();
            String NOMBRE_NODO_B = dt.Rows[0]["NOMBRE_NODO_B"].ToString();
            String TIPO_NODO_A = dt.Rows[0]["TIPO_NODO_A"].ToString();
            String TIPO_NODO_B = dt.Rows[0]["TIPO_NODO_B"].ToString();
            //String FRECUENCIA = ds.Tables[0].Rows[0]["FRECUENCIA"].ToString();

            String fechaSQL = dt.Rows[0]["FECHA"].ToString();
            String FECHA = "";
            if (fechaSQL != "")
            {
                DateTime dtFecha = DateTime.Parse(fechaSQL);
                FECHA = dtFecha.ToString("dd/MM/yyyy");
            }
            else { FECHA = ""; }

            #endregion

            #region Configuracion y Mediciones

            String DIREC_ESTACION_A = dt.Rows[0]["DIREC_ESTACION_A"].ToString();
            String DIREC_ESTACION_B = dt.Rows[0]["DIREC_ESTACION_B"].ToString();
            String DISTRITO_A = dt.Rows[0]["DISTRITO_A"].ToString();
            String DISTRITO_B = dt.Rows[0]["DISTRITO_B"].ToString();
            String PROVINCIA_A = dt.Rows[0]["PROVINCIA_A"].ToString();
            String PROVINCIA_B = dt.Rows[0]["PROVINCIA_B"].ToString();
            String DEPARTAMENTO_A = dt.Rows[0]["DEPARTAMENTO_A"].ToString();
            String DEPARTAMENTO_B = dt.Rows[0]["DEPARTAMENTO_B"].ToString();
            String LATITUD_A = dt.Rows[0]["LATITUD_A"].ToString();
            String LATITUD_B = dt.Rows[0]["LATITUD_B"].ToString();
            String LONGITUD_A = dt.Rows[0]["LONGITUD_A"].ToString();
            String LONGITUD_B = dt.Rows[0]["LONGITUD_B"].ToString();
            String REF_UBIC_EST_A = dt.Rows[0]["REF_UBIC_EST_A"].ToString();
            String REF_UBIC_EST_B = dt.Rows[0]["REF_UBIC_EST_B"].ToString();
            String ALTURA_TORRE_A = dt.Rows[0]["ALTURA_TORRE_A"].ToString();
            String ALTURA_TORRE_B = dt.Rows[0]["ALTURA_TORRE_B"].ToString();
            String DISTANCIA_A_B = dt.Rows[0]["DISTANCIA_A_B"].ToString();
            String MODULACION = dt.Rows[0]["MODULACION"].ToString();
            String CANAL_ESTAC_A = dt.Rows[0]["CANAL_ESTAC_A"].ToString();
            String CANAL_ESTAC_B = dt.Rows[0]["CANAL_ESTAC_B"].ToString();
            String VELOCIDAD_HABILITADA = dt.Rows[0]["VELOCIDAD_HABILITADA"].ToString(); //ES LA MISMA PARA AMBOS NODOS
            String POLARIZACION_A = dt.Rows[0]["POLARIZACION_A"].ToString();
            String POLARIZACION_B = dt.Rows[0]["POLARIZACION_B"].ToString();
            String AZIMUT_A = dt.Rows[0]["AZIMUT_A"].ToString();
            String AZIMUT_B = dt.Rows[0]["AZIMUT_B"].ToString();
            String IP_NODO_A_1 = dt.Rows[0]["IP_NODO_A_1"].ToString();
            String IP_NODO_A_2 = dt.Rows[0]["IP_NODO_A_2"].ToString();
            String IP_NODO_A_3 = dt.Rows[0]["IP_NODO_A_3"].ToString();
            String IP_NODO_A_4 = dt.Rows[0]["IP_NODO_A_4"].ToString();
            String IP_NODO_B_1 = dt.Rows[0]["IP_NODO_B_1"].ToString();
            String IP_NODO_B_2 = dt.Rows[0]["IP_NODO_B_2"].ToString();
            String IP_NODO_B_3 = dt.Rows[0]["IP_NODO_B_3"].ToString();
            String IP_NODO_B_4 = dt.Rows[0]["IP_NODO_B_4"].ToString();
            String DEFAULT_GATE_AB_1 = dt.Rows[0]["DEFAULT_GATE_AB_1"].ToString();
            String DEFAULT_GATE_AB_2 = dt.Rows[0]["DEFAULT_GATE_AB_2"].ToString();
            String DEFAULT_GATE_AB_3 = dt.Rows[0]["DEFAULT_GATE_AB_3"].ToString();
            String DEFAULT_GATE_AB_4 = dt.Rows[0]["DEFAULT_GATE_AB_4"].ToString();
            String POTENCIA_A = dt.Rows[0]["POTENCIA_A"].ToString();
            String POTENCIA_B = dt.Rows[0]["POTENCIA_B"].ToString();
            String MARGEN_DES_A = dt.Rows[0]["MARGEN_DES_A"].ToString();
            String MARGEN_DES_B = dt.Rows[0]["MARGEN_DES_B"].ToString();
            String NIVEL_UMBRAL_A_B = dt.Rows[0]["NIVEL_UMBRAL_A_B"].ToString();
            String NIVEL_RECEP_RADIO_A = dt.Rows[0]["NIVEL_RECEP_RADIO_A"].ToString();
            String NIVEL_RECEP_RADIO_B = dt.Rows[0]["NIVEL_RECEP_RADIO_B"].ToString();
            String PING_PTP_RADIO_A = dt.Rows[0]["PING_PTP_RADIO_A"].ToString();
            String PING_PTP_RADIO_B = dt.Rows[0]["PING_PTP_RADIO_B"].ToString();

            byte[] CONF_GEN_ENL_EST_A = (byte[])dt.Rows[0]["CONF_GEN_ENL_EST_A"];
            MemoryStream mCONF_GEN_ENL_EST_A = new MemoryStream(CONF_GEN_ENL_EST_A);
            byte[] CONF_GEN_ENL_EST_B = (byte[])dt.Rows[0]["CONF_GEN_ENL_EST_B"];
            MemoryStream mCONF_GEN_ENL_EST_B = new MemoryStream(CONF_GEN_ENL_EST_B);
            byte[] CONF_VLAN_ESTA_A = (byte[])dt.Rows[0]["CONF_VLAN_ESTA_A"];
            MemoryStream mCONF_VLAN_ESTA_A = new MemoryStream(CONF_VLAN_ESTA_A);
            byte[] CONF_VLAN_ESTA_B = (byte[])dt.Rows[0]["CONF_VLAN_ESTA_B"];
            MemoryStream mCONF_VLAN_ESTA_B = new MemoryStream(CONF_VLAN_ESTA_B);
            byte[] CONF_RADIO_ESTAC_A_1 = (byte[])dt.Rows[0]["CONF_RADIO_ESTAC_A_1"];
            MemoryStream mCONF_RADIO_ESTAC_A_1 = new MemoryStream(CONF_RADIO_ESTAC_A_1);
            byte[] CONF_RADIO_ESTAC_A_2 = (byte[])dt.Rows[0]["CONF_RADIO_ESTAC_A_2"];
            MemoryStream mCONF_RADIO_ESTAC_A_2 = new MemoryStream(CONF_RADIO_ESTAC_A_2);
            byte[] CONF_RADIO_ESTAC_B = (byte[])dt.Rows[0]["CONF_RADIO_ESTAC_B"];
            MemoryStream mCONF_RADIO_ESTAC_B = new MemoryStream(CONF_RADIO_ESTAC_B);
            byte[] CONF_IP_ESTAC_A = (byte[])dt.Rows[0]["CONF_IP_ESTAC_A"];
            MemoryStream mCONF_IP_ESTAC_A = new MemoryStream(CONF_IP_ESTAC_A);
            byte[] CONF_IP_ESTAC_B = (byte[])dt.Rows[0]["CONF_IP_ESTAC_B"];
            MemoryStream mCONF_IP_ESTAC_B = new MemoryStream(CONF_IP_ESTAC_B);

            #endregion

            #region Longitud SFTP

            //String VALOR_B_ESTAC_A = dt.Rows[0]["VALOR_B_ESTAC_A"].ToString();
            //String VALOR_C_ESTAC_A = dt.Rows[0]["VALOR_C_ESTAC_A"].ToString();
            //String VALOR_D_ESTAC_A = dt.Rows[0]["VALOR_D_ESTAC_A"].ToString();
            //String VALOR_E_ESTAC_A = dt.Rows[0]["VALOR_E_ESTAC_A"].ToString();
            //String VALOR_B_ESTAC_B = dt.Rows[0]["VALOR_B_ESTAC_B"].ToString();
            //String VALOR_C_ESTAC_B = dt.Rows[0]["VALOR_C_ESTAC_B"].ToString();
            //String VALOR_D_ESTAC_B = dt.Rows[0]["VALOR_D_ESTAC_B"].ToString();
            //String VALOR_E_ESTAC_B = dt.Rows[0]["VALOR_E_ESTAC_B"].ToString();

            #endregion

            #region Asignaciones y Observaciones

            String SWITCH_ROUTER_A = dt.Rows[0]["SWITCH_ROUTER_A"].ToString();
            String SWITCH_ROUTER_B = dt.Rows[0]["SWITCH_ROUTER_B"].ToString();
            String CAP_BREAKER_ASIG_EST_A = dt.Rows[0]["CAP_BREAKER_ASIG_EST_A"].ToString();
            String VOLT_DC_ESTAC_A = dt.Rows[0]["VOLT_DC_ESTAC_A"].ToString();
            String POS_BREAKER_ASIG_ESTAC_A = dt.Rows[0]["POS_BREAKER_ASIG_ESTAC_A"].ToString();
            String POS_BARRA_ATERRA_ESTA_A = dt.Rows[0]["POS_BARRA_ATERRA_ESTA_A"].ToString();
            String CAP_BREAKER_ASIG_EST_B = dt.Rows[0]["CAP_BREAKER_ASIG_EST_B"].ToString();
            String VOLT_DC_ESTAC_B = dt.Rows[0]["VOLT_DC_ESTAC_B"].ToString();
            String POS_BREAKER_ASIG_ESTAC_B = dt.Rows[0]["POS_BREAKER_ASIG_ESTAC_B"].ToString();
            String POS_BARRA_ATERRA_ESTA_B = dt.Rows[0]["POS_BARRA_ATERRA_ESTA_B"].ToString();

            #endregion

            #region Calculo Propagacion

            byte[] INGENIERIA = (byte[])dt.Rows[0]["INGENIERIA"];
            MemoryStream mINGENIERIA = new MemoryStream(INGENIERIA);
            byte[] PERFIL = (byte[])dt.Rows[0]["PERFIL"];
            MemoryStream mPERFIL = new MemoryStream(PERFIL);

            #endregion

            #region Pruebas de Interferencia

            byte[] PANT_RADIO_ESTAC_A = (byte[])dt.Rows[0]["PANT_RADIO_ESTAC_A"];
            MemoryStream mPANT_RADIO_ESTAC_A = new MemoryStream(PANT_RADIO_ESTAC_A);
            byte[] PANT_RADIO_ESTAC_B = (byte[])dt.Rows[0]["PANT_RADIO_ESTAC_B"];
            MemoryStream mPANT_RADIO_ESTAC_B = new MemoryStream(PANT_RADIO_ESTAC_B);

            #endregion

            #region Serie Equipos Fotos

            byte[] SERIE_ANT_ESTAC_A = (byte[])dt.Rows[0]["SERIE_ANT_ESTAC_A"];
            MemoryStream mSERIE_ANT_ESTAC_A = new MemoryStream(SERIE_ANT_ESTAC_A);
            byte[] SERIE_ODU_ESTAC_A = (byte[])dt.Rows[0]["SERIE_ODU_ESTAC_A"];
            MemoryStream mSERIE_ODU_ESTAC_A = new MemoryStream(SERIE_ODU_ESTAC_A);
            byte[] SERIE_POE_ESTAC_A = (byte[])dt.Rows[0]["SERIE_POE_ESTAC_A"];
            MemoryStream mSERIE_POE_ESTAC_A = new MemoryStream(SERIE_POE_ESTAC_A);

            CMM4BE CMM4A = new CMM4BE();
            List<CMM4BE> lstCMM4A = new List<CMM4BE>();
            CMM4A.Nodo.IdNodo = NODO_A;
            lstCMM4A = CMM4BL.ListarCMM4(CMM4A);


            CMM4BE CMM4B = new CMM4BE();
            List<CMM4BE> lstCMM4B = new List<CMM4BE>();
            CMM4A.Nodo.IdNodo = NODO_B;
            lstCMM4A = CMM4BL.ListarCMM4(CMM4B);


            byte[] SERIE_ANT_ESTAC_B = (byte[])dt.Rows[0]["SERIE_ANT_ESTAC_B"];
            MemoryStream mSERIE_ANT_ESTAC_B = new MemoryStream(SERIE_ANT_ESTAC_B);
            byte[] SERIE_ODU_ESTAC_B = (byte[])dt.Rows[0]["SERIE_ODU_ESTAC_B"];
            MemoryStream mSERIE_ODU_ESTAC_B = new MemoryStream(SERIE_ODU_ESTAC_B);
            byte[] SERIE_POE_ESTAC_B = (byte[])dt.Rows[0]["SERIE_POE_ESTAC_B"];
            MemoryStream mSERIE_POE_ESTAC_B = new MemoryStream(SERIE_POE_ESTAC_B);


            #endregion

            #region Reporte Fotografico

            #region Estacion A
            byte[] FOTO1_PAN_ESTAC_A = (byte[])dt.Rows[0]["FOTO1_PAN_ESTAC_A"];
            MemoryStream mFOTO1_PAN_ESTAC_A = new MemoryStream(FOTO1_PAN_ESTAC_A);
            byte[] FOTO2_POS_ANT_INST_TORRE_A = (byte[])dt.Rows[0]["FOTO2_POS_ANT_INST_TORRE_A"];
            MemoryStream mFOTO2_POS_ANT_INST_TORRE_A = new MemoryStream(FOTO2_POS_ANT_INST_TORRE_A);
            byte[] FOTO3_FOTO_ANTENA_PTP_ESTAC_A = (byte[])dt.Rows[0]["FOTO3_FOTO_ANTENA_PTP_ESTAC_A"];
            MemoryStream mFOTO3_FOTO_ANTENA_PTP_ESTAC_A = new MemoryStream(FOTO3_FOTO_ANTENA_PTP_ESTAC_A);
            byte[] FOTO4_ETIQ_PUERTO_ANT_ESTAC_A = (byte[])dt.Rows[0]["FOTO4_ETIQ_PUERTO_ANT_ESTAC_A"];
            MemoryStream mFOTO4_ETIQ_PUERTO_ANT_ESTAC_A = new MemoryStream(FOTO4_ETIQ_PUERTO_ANT_ESTAC_A);
            byte[] FOTO5_ENGRAS_PERNOS_ESTAC_A = (byte[])dt.Rows[0]["FOTO5_ENGRAS_PERNOS_ESTAC_A"];
            MemoryStream mFOTO5_ENGRAS_PERNOS_ESTAC_A = new MemoryStream(FOTO5_ENGRAS_PERNOS_ESTAC_A);
            byte[] FOTO6_SILC_CONECT_ESTAC_A = (byte[])dt.Rows[0]["FOTO6_SILC_CONECT_ESTAC_A"];
            MemoryStream mFOTO6_SILC_CONECT_ESTAC_A = new MemoryStream(FOTO6_SILC_CONECT_ESTAC_A);
            byte[] FOTO7_ATERRAM_ODU_TORRE_ESTAC_A = (byte[])dt.Rows[0]["FOTO7_ATERRAM_ODU_TORRE_ESTAC_A"];
            MemoryStream mFOTO7_ATERRAM_ODU_TORRE_ESTAC_A = new MemoryStream(FOTO7_ATERRAM_ODU_TORRE_ESTAC_A);
            byte[] FOTO8_RECORRI_SFTP_ESTAC_A = (byte[])dt.Rows[0]["FOTO8_RECORRI_SFTP_ESTAC_A"];
            MemoryStream mFOTO8_RECORRI_SFTP_ESTAC_A = new MemoryStream(FOTO8_RECORRI_SFTP_ESTAC_A);
            byte[] FOTO9_1__SFTP_OUT_1_ESTAC_A = (byte[])dt.Rows[0]["FOTO9_1__SFTP_OUT_1_ESTAC_A"];
            MemoryStream mFOTO9_1__SFTP_OUT_1_ESTAC_A = new MemoryStream(FOTO9_1__SFTP_OUT_1_ESTAC_A);
            byte[] FOTO9_2_SFTP_OUT_2_ESTAC_A = (byte[])dt.Rows[0]["FOTO9_2_SFTP_OUT_2_ESTAC_A"];
            MemoryStream mFOTO9_2_SFTP_OUT_2_ESTAC_A = new MemoryStream(FOTO9_2_SFTP_OUT_2_ESTAC_A);
            byte[] FOTO10_SALAN_OUT_ETIQ_ESTAC_A = (byte[])dt.Rows[0]["FOTO10_SALAN_OUT_ETIQ_ESTAC_A"];
            MemoryStream mFOTO10_SALAN_OUT_ETIQ_ESTAC_A = new MemoryStream(FOTO10_SALAN_OUT_ETIQ_ESTAC_A);
            byte[] FOTO11_ATERRAM_SALAN_ESTAC_A = (byte[])dt.Rows[0]["FOTO11_ATERRAM_SALAN_ESTAC_A"];
            MemoryStream mFOTO11_ATERRAM_SALAN_ESTAC_A = new MemoryStream(FOTO11_ATERRAM_SALAN_ESTAC_A);
            byte[] FOTO14_1_SFTP_IN_1_ESTAC_A = (byte[])dt.Rows[0]["FOTO14_1_SFTP_IN_1_ESTAC_A"];
            MemoryStream mFOTO14_1_SFTP_IN_1_ESTAC_A = new MemoryStream(FOTO14_1_SFTP_IN_1_ESTAC_A);
            byte[] FOTO14_2_SFTP_IN_2_ESTAC_A = (byte[])dt.Rows[0]["FOTO14_2_SFTP_IN_2_ESTAC_A"];
            MemoryStream mFOTO14_2_SFTP_IN_2_ESTAC_A = new MemoryStream(FOTO14_2_SFTP_IN_2_ESTAC_A);
            byte[] FOTO17_FOTO_PAN_RACK_ESTAC_A = (byte[])dt.Rows[0]["FOTO17_FOTO_PAN_RACK_ESTAC_A"];
            MemoryStream mFOTO17_FOTO_PAN_RACK_ESTAC_A = new MemoryStream(FOTO17_FOTO_PAN_RACK_ESTAC_A);
            byte[] FOTO18_1_ATERRAM_POE_ESTAC_A = (byte[])dt.Rows[0]["FOTO18_1_ATERRAM_POE_ESTAC_A"];
            MemoryStream mFOTO18_1_ATERRAM_POE_ESTAC_A = new MemoryStream(FOTO18_1_ATERRAM_POE_ESTAC_A);
            byte[] FOTO18_2_ATERRAM_POE_ESTAC_A = (byte[])dt.Rows[0]["FOTO18_2_ATERRAM_POE_ESTAC_A"];
            MemoryStream mFOTO18_2_ATERRAM_POE_ESTAC_A = new MemoryStream(FOTO18_2_ATERRAM_POE_ESTAC_A);
            byte[] FOTO19_1_ENERG_POE_ETIQ_ESTAC_A = (byte[])dt.Rows[0]["FOTO19_1_ENERG_POE_ETIQ_ESTAC_A"];
            MemoryStream mFOTO19_1_ENERG_POE_ETIQ_ESTAC_A = new MemoryStream(FOTO19_1_ENERG_POE_ETIQ_ESTAC_A);
            byte[] FOTO19_2_ENERG_POE_ETIQ_ESTAC_A = (byte[])dt.Rows[0]["FOTO19_2_ENERG_POE_ETIQ_ESTAC_A"];
            MemoryStream mFOTO19_2_ENERG_POE_ETIQ_ESTAC_A = new MemoryStream(FOTO19_2_ENERG_POE_ETIQ_ESTAC_A);
            byte[] FOTO20_1_PATCH_POE_ESTAC_A = (byte[])dt.Rows[0]["FOTO20_1_PATCH_POE_ESTAC_A"];
            MemoryStream mFOTO20_1_PATCH_POE_ESTAC_A = new MemoryStream(FOTO20_1_PATCH_POE_ESTAC_A);
            byte[] FOTO20_2_PATCH_POE_ESTAC_A = (byte[])dt.Rows[0]["FOTO20_2_PATCH_POE_ESTAC_A"];
            MemoryStream mFOTO20_2_PATCH_POE_ESTAC_A = new MemoryStream(FOTO20_2_PATCH_POE_ESTAC_A);
            #endregion

            #region Estacion B

            //byte[] FOTO1_PAN_ESTAC_B = (byte[])ds.Tables[0].Rows[0]["FOTO1_PAN_ESTAC_B"];
            //MemoryStream mFOTO1_PAN_ESTAC_B = new MemoryStream(FOTO1_PAN_ESTAC_B);
            //byte[] FOTO2_POS_ANT_INST_TORRE_B = (byte[])ds.Tables[0].Rows[0]["FOTO2_POS_ANT_INST_TORRE_B"];
            //MemoryStream mFOTO2_POS_ANT_INST_TORRE_B = new MemoryStream(FOTO2_POS_ANT_INST_TORRE_B);
            //byte[] FOTO3_FOTO_ANTENA_PTP_ESTAC_B = (byte[])ds.Tables[0].Rows[0]["FOTO3_FOTO_ANTENA_PTP_ESTAC_B"];
            //MemoryStream mFOTO3_FOTO_ANTENA_PTP_ESTAC_B = new MemoryStream(FOTO3_FOTO_ANTENA_PTP_ESTAC_B);
            //byte[] FOTO4_ETIQ_PUERTO_ANT_ESTAC_B = (byte[])ds.Tables[0].Rows[0]["FOTO4_ETIQ_PUERTO_ANT_ESTAC_B"];
            //MemoryStream mFOTO4_ETIQ_PUERTO_ANT_ESTAC_B = new MemoryStream(FOTO4_ETIQ_PUERTO_ANT_ESTAC_B);
            //byte[] FOTO5_ENGRAS_PERNOS_ESTAC_B = (byte[])ds.Tables[0].Rows[0]["FOTO5_ENGRAS_PERNOS_ESTAC_B"];
            //MemoryStream mFOTO5_ENGRAS_PERNOS_ESTAC_B = new MemoryStream(FOTO5_ENGRAS_PERNOS_ESTAC_B);
            //byte[] FOTO6_SILC_CONECT_ESTAC_B = (byte[])ds.Tables[0].Rows[0]["FOTO6_SILC_CONECT_ESTAC_B"];
            //MemoryStream mFOTO6_SILC_CONECT_ESTAC_B = new MemoryStream(FOTO6_SILC_CONECT_ESTAC_B);
            //byte[] FOTO7_ATERRAM_ODU_TORRE_ESTAC_B = (byte[])ds.Tables[0].Rows[0]["FOTO7_ATERRAM_ODU_TORRE_ESTAC_B"];
            //MemoryStream mFOTO7_ATERRAM_ODU_TORRE_ESTAC_B = new MemoryStream(FOTO7_ATERRAM_ODU_TORRE_ESTAC_B);
            //byte[] FOTO8_RECORRI_SFTP_ESTAC_B = (byte[])ds.Tables[0].Rows[0]["FOTO8_RECORRI_SFTP_ESTAC_B"];
            //MemoryStream mFOTO8_RECORRI_SFTP_ESTAC_B = new MemoryStream(FOTO8_RECORRI_SFTP_ESTAC_B);
            //byte[] FOTO9_1__SFTP_OUT_1_ESTAC_B = (byte[])ds.Tables[0].Rows[0]["FOTO9_1__SFTP_OUT_1_ESTAC_B"];
            //MemoryStream mFOTO9_1__SFTP_OUT_1_ESTAC_B = new MemoryStream(FOTO9_1__SFTP_OUT_1_ESTAC_B);
            //byte[] FOTO9_2_SFTP_OUT_2_ESTAC_B = (byte[])ds.Tables[0].Rows[0]["FOTO9_2_SFTP_OUT_2_ESTAC_B"];
            //MemoryStream mFOTO9_2_SFTP_OUT_2_ESTAC_B = new MemoryStream(FOTO9_2_SFTP_OUT_2_ESTAC_B);
            //byte[] FOTO10_SALAN_OUT_ETIQ_ESTAC_B = (byte[])ds.Tables[0].Rows[0]["FOTO10_SALAN_OUT_ETIQ_ESTAC_B"];
            //MemoryStream mFOTO10_SALAN_OUT_ETIQ_ESTAC_B = new MemoryStream(FOTO10_SALAN_OUT_ETIQ_ESTAC_B);
            //byte[] FOTO11_ATERRAM_SALAN_ESTAC_B = (byte[])ds.Tables[0].Rows[0]["FOTO11_ATERRAM_SALAN_ESTAC_B"];
            //MemoryStream mFOTO11_ATERRAM_SALAN_ESTAC_B = new MemoryStream(FOTO11_ATERRAM_SALAN_ESTAC_B);
            //byte[] FOTO14_1_SFTP_IN_1_ESTAC_B = (byte[])ds.Tables[0].Rows[0]["FOTO14_1_SFTP_IN_1_ESTAC_B"];
            //MemoryStream mFOTO14_1_SFTP_IN_1_ESTAC_B = new MemoryStream(FOTO14_1_SFTP_IN_1_ESTAC_B);
            //byte[] FOTO14_2_SFTP_IN_2_ESTAC_B = (byte[])ds.Tables[0].Rows[0]["FOTO14_2_SFTP_IN_2_ESTAC_B"];
            //MemoryStream mFOTO14_2_SFTP_IN_2_ESTAC_B = new MemoryStream(FOTO14_2_SFTP_IN_2_ESTAC_B);
            //byte[] FOTO17_FOTO_PAN_RACK_ESTAC_B = (byte[])ds.Tables[0].Rows[0]["FOTO17_FOTO_PAN_RACK_ESTAC_B"];
            //MemoryStream mFOTO17_FOTO_PAN_RACK_ESTAC_B = new MemoryStream(FOTO17_FOTO_PAN_RACK_ESTAC_B);
            //byte[] FOTO18_1_ATERRAM_POE_ESTAC_B = (byte[])ds.Tables[0].Rows[0]["FOTO18_1_ATERRAM_POE_ESTAC_B"];
            //MemoryStream mFOTO18_1_ATERRAM_POE_ESTAC_B = new MemoryStream(FOTO18_1_ATERRAM_POE_ESTAC_B);
            //byte[] FOTO18_2_ATERRAM_POE_ESTAC_B = (byte[])ds.Tables[0].Rows[0]["FOTO18_2_ATERRAM_POE_ESTAC_B"];
            //MemoryStream mFOTO18_2_ATERRAM_POE_ESTAC_B = new MemoryStream(FOTO18_2_ATERRAM_POE_ESTAC_B);
            //byte[] FOTO19_1_ENERG_POE_ETIQ_ESTAC_B = (byte[])ds.Tables[0].Rows[0]["FOTO19_1_ENERG_POE_ETIQ_ESTAC_B"];
            //MemoryStream mFOTO19_1_ENERG_POE_ETIQ_ESTAC_B = new MemoryStream(FOTO19_1_ENERG_POE_ETIQ_ESTAC_B);
            //byte[] FOTO19_2_ENERG_POE_ETIQ_ESTAC_B = (byte[])ds.Tables[0].Rows[0]["FOTO19_2_ENERG_POE_ETIQ_ESTAC_B"];
            //MemoryStream mFOTO19_2_ENERG_POE_ETIQ_ESTAC_B = new MemoryStream(FOTO19_2_ENERG_POE_ETIQ_ESTAC_B);
            //byte[] FOTO20_1_PATCH_POE_ESTAC_B = (byte[])ds.Tables[0].Rows[0]["FOTO20_1_PATCH_POE_ESTAC_B"];
            //MemoryStream mFOTO20_1_PATCH_POE_ESTAC_B = new MemoryStream(FOTO20_1_PATCH_POE_ESTAC_B);
            //byte[] FOTO20_2_PATCH_POE_ESTAC_B = (byte[])ds.Tables[0].Rows[0]["FOTO20_2_PATCH_POE_ESTAC_B"];
            //MemoryStream mFOTO20_2_PATCH_POE_ESTAC_B = new MemoryStream(FOTO20_2_PATCH_POE_ESTAC_B);


            #endregion

            #endregion

            #region Datos Generales Nodo A

            String UBIGEO_A = dt.Rows[0]["UBIGEO_A"].ToString();
            String SERIE_PTP_450_A = dt.Rows[0]["SERIE_PTP_450_A"].ToString();
            String FRECUENCIA_A = dt.Rows[0]["FRECUENCIA_A"].ToString();
            String PIRE_EIRP_A = dt.Rows[0]["PIRE_EIRP_A"].ToString();
            String ANTENA_MARCA_MODELO_A = dt.Rows[0]["ANTENA_MARCA_MODELO_A"].ToString();
            String GANANCIA_ANTENA_A = dt.Rows[0]["GANANCIA_ANTENA_A"].ToString();
            String ALTURA_ANTENA_A = dt.Rows[0]["ALTURA_ANTENA_A"].ToString();
            String ELEVACION_A = dt.Rows[0]["ELEVACION_A"].ToString();
            String ALTITUD_msnm_A = dt.Rows[0]["ALTITUD_msnm_A"].ToString();

            #endregion

            #region Datos Generales Nodo B

            String UBIGEO_B = dt.Rows[0]["UBIGEO_B"].ToString();
            String SERIE_PTP_450_B = dt.Rows[0]["SERIE_PTP_450_B"].ToString();
            String FRECUENCIA_B = dt.Rows[0]["FRECUENCIA_B"].ToString();
            String PIRE_EIRP_B = dt.Rows[0]["PIRE_EIRP_B"].ToString();
            String ANTENA_MARCA_MODELO_B = dt.Rows[0]["ANTENA_MARCA_MODELO_B"].ToString();
            String GANANCIA_ANTENA_B = dt.Rows[0]["GANANCIA_ANTENA_B"].ToString();
            String ALTURA_ANTENA_B = dt.Rows[0]["ALTURA_ANTENA_B"].ToString();
            String ELEVACION_B = dt.Rows[0]["ELEVACION_B"].ToString();
            String ALTITUD_msnm_B = dt.Rows[0]["ALTITUD_msnm_B"].ToString();

            #endregion

            #endregion

            // String usuarioWindows = Environment.UserName;
            String excelGenerado = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";

            File.Copy(rutaPlantilla, excelGenerado, true);

            #region Ingresando Valores

            #region Caratula

            ExcelToolsBL.UpdateCell(excelGenerado, "Carátula", NOMBRE_NODO_A, 15, "C");
            ExcelToolsBL.UpdateCell(excelGenerado, "Carátula", NODO_A, 18, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "Carátula", NODO_B, 18, "F");
            ExcelToolsBL.UpdateCell(excelGenerado, "Carátula", NOMBRE_NODO_A, 19, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "Carátula", NOMBRE_NODO_A, 19, "F");
            ExcelToolsBL.UpdateCell(excelGenerado, "Carátula", TIPO_NODO_A, 20, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "Carátula", TIPO_NODO_B, 20, "F");
            ExcelToolsBL.UpdateCell(excelGenerado, "Carátula", FRECUENCIA_A, 22, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "Carátula", FECHA, 24, "E");

            #endregion

            #region Configuracion y Mediciones

            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", TIPO_NODO_A, 16, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", TIPO_NODO_B, 16, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", NODO_A, 17, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", NODO_B, 17, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", NOMBRE_NODO_A, 18, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", NOMBRE_NODO_B, 18, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", DIREC_ESTACION_A, 19, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", DIREC_ESTACION_B, 19, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", DISTRITO_A, 20, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", DISTRITO_B, 20, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", PROVINCIA_A, 21, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", PROVINCIA_B, 21, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", DEPARTAMENTO_A, 22, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", DEPARTAMENTO_B, 22, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", LATITUD_A, 23, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", LATITUD_B, 23, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", LONGITUD_A, 24, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", LONGITUD_B, 24, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", ALTITUD_msnm_A, 25, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", ALTITUD_msnm_B, 25, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", REF_UBIC_EST_A, 26, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", REF_UBIC_EST_B, 26, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", ALTURA_TORRE_A, 28, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", ALTURA_TORRE_B, 28, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", DISTANCIA_A_B, 29, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", MODULACION, 31, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", MODULACION, 31, "j");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", FRECUENCIA_A, 32, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", FRECUENCIA_B, 32, "j");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", FRECUENCIA_A, 33, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", FRECUENCIA_B, 33, "j");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", CANAL_ESTAC_A, 34, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", CANAL_ESTAC_B, 34, "j");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", VELOCIDAD_HABILITADA, 37, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", VELOCIDAD_HABILITADA, 37, "j");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", AZIMUT_A, 43, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", AZIMUT_B, 43, "j");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", "Modulación:" + MODULACION + "QAM", 50, "K");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", IP_NODO_A_1, 55, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", IP_NODO_A_2, 55, "F");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", IP_NODO_A_3, 55, "G");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", IP_NODO_A_4, 55, "H");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", IP_NODO_B_1, 55, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", IP_NODO_B_2, 55, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", IP_NODO_B_3, 55, "K");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", IP_NODO_B_4, 55, "L");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", IP_NODO_A_1, 56, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", IP_NODO_A_2, 56, "F");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", IP_NODO_A_3, 56, "G");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", IP_NODO_A_4, 56, "H");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", IP_NODO_B_1, 56, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", IP_NODO_B_2, 56, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", IP_NODO_B_3, 56, "K");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", IP_NODO_B_4, 56, "L");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", DEFAULT_GATE_AB_1, 58, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", DEFAULT_GATE_AB_2, 58, "F");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", DEFAULT_GATE_AB_3, 58, "G");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", DEFAULT_GATE_AB_4, 58, "H");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", DEFAULT_GATE_AB_1, 58, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", DEFAULT_GATE_AB_2, 58, "J");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", DEFAULT_GATE_AB_3, 58, "K");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", DEFAULT_GATE_AB_4, 58, "L");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", POTENCIA_A, 71, "K");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", POTENCIA_B, 71, "M");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", POTENCIA_A, 72, "K");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", POTENCIA_B, 72, "M");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", MARGEN_DES_A, 73, "K");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", MARGEN_DES_B, 73, "M");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", NIVEL_UMBRAL_A_B, 74, "K");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", NIVEL_UMBRAL_A_B, 74, "M");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", NIVEL_RECEP_RADIO_A, 78, "K");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", NIVEL_RECEP_RADIO_B, 78, "M");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", PING_PTP_RADIO_A + " ms", 87, "K");
            ExcelToolsBL.UpdateCell(excelGenerado, "1 Configuración y Mediciones", PING_PTP_RADIO_B + " ms", 87, "M");

            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Mediciones", mCONF_GEN_ENL_EST_A, "", 94, 3, 712, 334);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Mediciones", mCONF_GEN_ENL_EST_B, "", 110, 3, 709, 299);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Mediciones", mCONF_VLAN_ESTA_A, "", 128, 3, 710, 285);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Mediciones", mCONF_VLAN_ESTA_B, "", 147, 3, 710, 284);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Mediciones", mCONF_RADIO_ESTAC_A_1, "", 170, 3, 334, 301);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Mediciones", mCONF_RADIO_ESTAC_A_2, "", 170, 7, 380, 298);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Mediciones", mCONF_RADIO_ESTAC_B, "", 187, 3, 688, 587);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Mediciones", mCONF_IP_ESTAC_A, "", 229, 3, 710, 231);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "1 Configuración y Mediciones", mCONF_IP_ESTAC_B, "", 245, 3, 687, 328);

            #endregion

            #region Materiales A

            ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales A", NOMBRE_NODO_A + " - " + NOMBRE_NODO_B, 12, "C");
            ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales A", NOMBRE_NODO_A, 12, "F");

            foreach (DataRow dr in dt1.Rows)
            {
                String EQUIPO = dr["EQUIPO"].ToString();
                String MARCA = dr["MARCA"].ToString();
                String MODELO = dr["MODELO"].ToString();
                String Nro_SERIE = dr["Nro_SERIE"].ToString();

                int ind = Convert.ToInt32(dt1.Rows.IndexOf(dr));

                ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales A", Convert.ToString(ind + 1), 17 + ind, "B");
                ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales A", EQUIPO, 17 + ind, "C");
                ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales A", "1", 17 + ind, "D");
                ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales A", MARCA, 17 + ind, "E");
                ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales A", MODELO, 17 + ind, "F");
                ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales A", Nro_SERIE, 17 + ind, "G");

            }

            foreach (DataRow dr in dt2.Rows)
            {
                String DESCRIPCION = dr["DESCRIPCION"].ToString();
                String UNIDAD = dr["UNIDAD"].ToString();
                String CANTIDAD = dr["CANTIDAD"].ToString();

                int ind = Convert.ToInt32(dt2.Rows.IndexOf(dr));

                ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales A", Convert.ToString(ind + 1), 32 + ind, "B");
                ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales A", DESCRIPCION, 32 + ind, "C");
                ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales A", UNIDAD, 32 + ind, "E");
                ExcelToolsBL.UpdateCell(excelGenerado, "2 Materiales A", CANTIDAD, 32 + ind, "F");

            }

            #endregion

            #region Materiales B

            ExcelToolsBL.UpdateCell(excelGenerado, "3 Materiales B", NOMBRE_NODO_A + " - " + NOMBRE_NODO_B, 12, "C");
            ExcelToolsBL.UpdateCell(excelGenerado, "3 Materiales B", NOMBRE_NODO_B, 12, "F");

            foreach (DataRow dr in dt3.Rows)
            {
                String EQUIPO = dr["EQUIPO"].ToString();
                String MARCA = dr["MARCA"].ToString();
                String MODELO = dr["MODELO"].ToString();
                String Nro_SERIE = dr["Nro_SERIE"].ToString();

                int ind = Convert.ToInt32(dt3.Rows.IndexOf(dr));

                ExcelToolsBL.UpdateCell(excelGenerado, "3 Materiales B", Convert.ToString(ind + 1), 17 + ind, "B");
                ExcelToolsBL.UpdateCell(excelGenerado, "3 Materiales B", EQUIPO, 17 + ind, "C");
                ExcelToolsBL.UpdateCell(excelGenerado, "3 Materiales B", "1", 17 + ind, "D");
                ExcelToolsBL.UpdateCell(excelGenerado, "3 Materiales B", MARCA, 17 + ind, "E");
                ExcelToolsBL.UpdateCell(excelGenerado, "3 Materiales B", MODELO, 17 + ind, "F");
                ExcelToolsBL.UpdateCell(excelGenerado, "3 Materiales B", Nro_SERIE, 17 + ind, "G");

            }

            foreach (DataRow dr in dt4.Rows)
            {
                String DESCRIPCION = dr["DESCRIPCION"].ToString();
                String UNIDAD = dr["UNIDAD"].ToString();
                String CANTIDAD = dr["CANTIDAD"].ToString();

                int ind = Convert.ToInt32(dt4.Rows.IndexOf(dr));

                ExcelToolsBL.UpdateCell(excelGenerado, "3 Materiales B", Convert.ToString(ind + 1), 32 + ind, "B");
                ExcelToolsBL.UpdateCell(excelGenerado, "3 Materiales B", DESCRIPCION, 32 + ind, "C");
                ExcelToolsBL.UpdateCell(excelGenerado, "3 Materiales B", UNIDAD, 32 + ind, "E");
                ExcelToolsBL.UpdateCell(excelGenerado, "3 Materiales B", CANTIDAD, 32 + ind, "F");

            }

            #endregion

            #region Longitud SFTP

            ExcelToolsBL.UpdateCell(excelGenerado, "4 Longitud SFTP", NOMBRE_NODO_A + " - " + NOMBRE_NODO_B, 13, "D");

            ExcelToolsBL.UpdateCell(excelGenerado, "4 Longitud SFTP", DIREC_ESTACION_A, 20, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "4 Longitud SFTP", ALTURA_TORRE_A, 20, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "4 Longitud SFTP", ALTURA_ANTENA_A, 20, "F");
            ExcelToolsBL.UpdateCell(excelGenerado, "4 Longitud SFTP", "3", 20, "G");
            ExcelToolsBL.UpdateCell(excelGenerado, "4 Longitud SFTP", "2,6", 20, "h");
            ExcelToolsBL.UpdateCell(excelGenerado, "4 Longitud SFTP", "8", 20, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "4 Longitud SFTP", "7", 20, "J");

            ExcelToolsBL.UpdateCell(excelGenerado, "4 Longitud SFTP", DIREC_ESTACION_B, 22, "D");
            ExcelToolsBL.UpdateCell(excelGenerado, "4 Longitud SFTP", ALTURA_TORRE_B, 22, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "4 Longitud SFTP", ALTURA_ANTENA_B, 22, "F");
            ExcelToolsBL.UpdateCell(excelGenerado, "4 Longitud SFTP", "3", 22, "G");
            ExcelToolsBL.UpdateCell(excelGenerado, "4 Longitud SFTP", "2,6", 22, "h");
            ExcelToolsBL.UpdateCell(excelGenerado, "4 Longitud SFTP", "8", 22, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "4 Longitud SFTP", "7", 22, "J");


            #endregion

            #region Asignaciones y Observaciones

            if (SWITCH_ROUTER_A.First().Equals("7"))
            {
                ExcelToolsBL.UpdateCell(excelGenerado, "5 Asignaciones y Observaciones", "Router Nokia Puerto" + " " + SWITCH_ROUTER_A.Last(), 17, "G");
            }
            else
            {
                ExcelToolsBL.UpdateCell(excelGenerado, "5 Asignaciones y Observaciones", "Omni Switch Puerto" + " " + SWITCH_ROUTER_A.Last(), 17, "G");
            }
            if (SWITCH_ROUTER_B.First().Equals("7"))
            {
                ExcelToolsBL.UpdateCell(excelGenerado, "5 Asignaciones y Observaciones", "Router Nokia Puerto" + " " + SWITCH_ROUTER_B.Last(), 22, "G");
            }
            else
            {
                ExcelToolsBL.UpdateCell(excelGenerado, "5 Asignaciones y Observaciones", "Omni Switch Puerto" + " " + SWITCH_ROUTER_B.Last(), 22, "G");
            }

            ExcelToolsBL.UpdateCell(excelGenerado, "5 Asignaciones y Observaciones", CAP_BREAKER_ASIG_EST_A, 27, "K");
            ExcelToolsBL.UpdateCell(excelGenerado, "5 Asignaciones y Observaciones", VOLT_DC_ESTAC_A, 28, "K");
            ExcelToolsBL.UpdateCell(excelGenerado, "5 Asignaciones y Observaciones", POS_BREAKER_ASIG_ESTAC_A, 29, "K");
            ExcelToolsBL.UpdateCell(excelGenerado, "5 Asignaciones y Observaciones", POS_BARRA_ATERRA_ESTA_A, 30, "K");
            ExcelToolsBL.UpdateCell(excelGenerado, "5 Asignaciones y Observaciones", CAP_BREAKER_ASIG_EST_B, 34, "K");
            ExcelToolsBL.UpdateCell(excelGenerado, "5 Asignaciones y Observaciones", VOLT_DC_ESTAC_B, 35, "K");
            ExcelToolsBL.UpdateCell(excelGenerado, "5 Asignaciones y Observaciones", POS_BREAKER_ASIG_ESTAC_B, 36, "K");
            ExcelToolsBL.UpdateCell(excelGenerado, "5 Asignaciones y Observaciones", POS_BARRA_ATERRA_ESTA_B, 37, "K");


            #endregion

            #region Calculo Propagacion

            ExcelToolsBL.AddImageDocument(false, excelGenerado, "6 Cálculo Propagacion", mINGENIERIA, "", 8, 3, 632, 528);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "6 Cálculo Propagacion", mPERFIL, "", 33, 2, 654, 368);

            #endregion

            #region Pruebas de Interferencia

            ExcelToolsBL.UpdateCell(excelGenerado, "8 Pruebas de Interferencia", NODO_A, 15, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "8 Pruebas de Interferencia", NODO_B, 15, "F");
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "8 Pruebas de Interferencia", mPANT_RADIO_ESTAC_A, "", 20, 3, 700, 268);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "8 Pruebas de Interferencia", mPANT_RADIO_ESTAC_B, "", 38, 3, 698, 269);

            #endregion

            #region Serie Equipos Fotos

            ExcelToolsBL.AddImageDocument(false, excelGenerado, "9 Serie Equipos (fotos)", mSERIE_ANT_ESTAC_A, "", 14, 3, 120, 98);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "9 Serie Equipos (fotos)", mSERIE_ODU_ESTAC_A, "", 21, 3, 81, 113);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "9 Serie Equipos (fotos)", mSERIE_POE_ESTAC_A, "", 29, 3, 152, 117);


            if (!lstCMM4A.Count.Equals(0))
            {
                byte[] SERIE_CMM4_ESTAC_A = (byte[])dt.Rows[0]["SERIE_CMM4_ESTAC_A"];
                MemoryStream mSERIE_CMM4_ESTAC_A = new MemoryStream(SERIE_CMM4_ESTAC_A);
                byte[] SERIE_UGPS_ESTAC_A = (byte[])dt.Rows[0]["SERIE_UGPS_ESTAC_A"];
                MemoryStream mSERIE_UGPS_ESTAC_A = new MemoryStream(SERIE_UGPS_ESTAC_A);
                byte[] SERIE_CONVERSOR_ESTAC_A = (byte[])dt.Rows[0]["SERIE_CONVERSOR_ESTAC_A"];
                MemoryStream mSERIE_CONVERSOR_ESTAC_A = new MemoryStream(SERIE_CONVERSOR_ESTAC_A);

                ExcelToolsBL.AddImageDocument(false, excelGenerado, "9 Serie Equipos (fotos)", mSERIE_CMM4_ESTAC_A, "", 37, 3, 152, 117);
                ExcelToolsBL.AddImageDocument(false, excelGenerado, "9 Serie Equipos (fotos)", mSERIE_UGPS_ESTAC_A, "", 45, 3, 152, 117);
                ExcelToolsBL.AddImageDocument(false, excelGenerado, "9 Serie Equipos (fotos)", mSERIE_CONVERSOR_ESTAC_A, "", 52, 3, 152, 117);

            }

            if (!lstCMM4B.Count.Equals(0))
            {
                byte[] SERIE_CMM4_ESTAC_B = (byte[])dt.Rows[0]["SERIE_CMM4_ESTAC_B"];
                MemoryStream mSERIE_CMM4_ESTAC_B = new MemoryStream(SERIE_CMM4_ESTAC_B);
                byte[] SERIE_UGPS_ESTAC_B = (byte[])dt.Rows[0]["SERIE_UGPS_ESTAC_B"];
                MemoryStream mSERIE_UGPS_ESTAC_B = new MemoryStream(SERIE_UGPS_ESTAC_B);
                byte[] SERIE_CONVERSOR_ESTAC_B = (byte[])dt.Rows[0]["SERIE_CONVERSOR_ESTAC_B"];
                MemoryStream mSERIE_CONVERSOR_ESTAC_B = new MemoryStream(SERIE_CONVERSOR_ESTAC_B);

                ExcelToolsBL.AddImageDocument(false, excelGenerado, "9 Serie Equipos (fotos)", mSERIE_CMM4_ESTAC_B, "", 82, 3, 152, 117);
                ExcelToolsBL.AddImageDocument(false, excelGenerado, "9 Serie Equipos (fotos)", mSERIE_UGPS_ESTAC_B, "", 90, 3, 152, 117);
                ExcelToolsBL.AddImageDocument(false, excelGenerado, "9 Serie Equipos (fotos)", mSERIE_CONVERSOR_ESTAC_B, "", 97, 3, 152, 117);

            }


            ExcelToolsBL.AddImageDocument(false, excelGenerado, "9 Serie Equipos (fotos)", mSERIE_ANT_ESTAC_B, "", 59, 3, 179, 23);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "9 Serie Equipos (fotos)", mSERIE_ODU_ESTAC_B, "", 66, 3, 81, 113);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "9 Serie Equipos (fotos)", mSERIE_POE_ESTAC_B, "", 75, 3, 152, 117);


            #endregion

            #region Reporte Fotografico

            #region Estacion A

            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO1_PAN_ESTAC_A, "",12,3, 408,383);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO2_POS_ANT_INST_TORRE_A,"", 12, 14,408, 383);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO3_FOTO_ANTENA_PTP_ESTAC_A, "", 29,3,407, 256);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO4_ETIQ_PUERTO_ANT_ESTAC_A, "", 29,14,408, 256);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO5_ENGRAS_PERNOS_ESTAC_A, "",46,3,407,289);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO6_SILC_CONECT_ESTAC_A, "", 46,14,408,291);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO7_ATERRAM_ODU_TORRE_ESTAC_A, "", 62,3,407, 467);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO8_RECORRI_SFTP_ESTAC_A, "",62,14,406,469);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO9_1__SFTP_OUT_1_ESTAC_A, "",79,3,408,188);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO9_2_SFTP_OUT_2_ESTAC_A, "", 90,3,408,164);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO10_SALAN_OUT_ETIQ_ESTAC_A, "",79,14,406, 350);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO11_ATERRAM_SALAN_ESTAC_A, "",101,3,408, 250);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO14_1_SFTP_IN_1_ESTAC_A, "",120,14,407, 188);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO14_2_SFTP_IN_2_ESTAC_A, "",131,14,407, 165);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO17_FOTO_PAN_RACK_ESTAC_A, "",152,3,410, 349);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO18_1_ATERRAM_POE_ESTAC_A, "",152,14,406, 184);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO18_2_ATERRAM_POE_ESTAC_A, "",163,14,406, 166);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO19_1_ENERG_POE_ETIQ_ESTAC_A, "",169,3,407, 187);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO19_2_ENERG_POE_ETIQ_ESTAC_A,"",180,3,409, 190);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO20_1_PATCH_POE_ESTAC_A, "", 169,14,407, 187);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO20_2_PATCH_POE_ESTAC_A,"",180,14,406,187);

            #endregion

            #region Estacion B

            //ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO1_PAN_ESTAC_A, "", 192, 3, 401, 368);
            //ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO2_POS_ANT_INST_TORRE_A, "", 192, 14, 415, 378);
            //ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO3_FOTO_ANTENA_PTP_ESTAC_A, "", 29, 3, 397, 239);
            //ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO4_ETIQ_PUERTO_ANT_ESTAC_A, "", 29, 16, 189, 250);
            //ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO5_ENGRAS_PERNOS_ESTAC_A, "", 46, 3, 362, 278);
            //ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO6_SILC_CONECT_ESTAC_A, "", 46, 14, 189, 250);
            //ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO7_ATERRAM_ODU_TORRE_ESTAC_A, "", 62, 3, 396, 448);
            //ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO8_RECORRI_SFTP_ESTAC_A, "", 62, 14, 393, 423);
            //ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO9_1__SFTP_OUT_1_ESTAC_A, "", 79, 4, 244, 176);
            //ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO9_2_SFTP_OUT_2_ESTAC_A, "", 90, 5, 244, 176);
            //ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO10_SALAN_OUT_ETIQ_ESTAC_A, "", 79, 15, 295, 327);
            //ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO11_ATERRAM_SALAN_ESTAC_A, "", 101, 3, 399, 245);
            //ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO14_1_SFTP_IN_1_ESTAC_A, "", 120, 14, 405, 185);
            //ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO14_2_SFTP_IN_2_ESTAC_A, "", 131, 14, 406, 157);
            //ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO17_FOTO_PAN_RACK_ESTAC_A, "", 152, 3, 377, 338);
            //ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO18_1_ATERRAM_POE_ESTAC_A, "", 152, 15, 334, 179);
            //ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO18_2_ATERRAM_POE_ESTAC_A, "", 163, 15, 264, 166);
            //ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO18_2_ATERRAM_POE_ESTAC_A, "", 163, 15, 264, 166);
            //ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO19_1_ENERG_POE_ETIQ_ESTAC_A, "", 169, 3, 345, 175);
            //ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO19_2_ENERG_POE_ETIQ_ESTAC_A, "", 180, 3, 385, 183);
            //ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO20_1_PATCH_POE_ESTAC_A, "", 169, 15, 296, 178);
            //ExcelToolsBL.AddImageDocument(false, excelGenerado, "10 Reporte Fotográfico", mFOTO20_2_PATCH_POE_ESTAC_A, "", 180, 14, 323, 162);

            #endregion


            #endregion

            #region Datos Generales Nodo A

            ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", TIPO_NODO_A, 9, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", NODO_A, 12, "B");
            ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", NOMBRE_NODO_A, 12, "G");
            ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", NOMBRE_NODO_A, 17, "B");
            ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", UBIGEO_A, 17, "G");
            ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", DEPARTAMENTO_A, 20, "B");
            ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", PROVINCIA_A, 20, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", DISTRITO_A, 20, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", SERIE_PTP_450_A, 29, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", FRECUENCIA_A + " Mhz", 40, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", PIRE_EIRP_A + " dbm", 44, "I");

            if (ANTENA_MARCA_MODELO_A.Equals("INTEGRADA"))
            {
                ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", "Cambium Networks/5092HH", 46, "I");

            }
            else if (ANTENA_MARCA_MODELO_A.Equals("0.6"))
            {

                ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", "SHENGLU/SLU0652DD6B", 46, "I");
            }

            ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", GANANCIA_ANTENA_A, 47, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", ALTURA_ANTENA_A, 49, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", ELEVACION_A, 52, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", ALTITUD_msnm_A, 53, "I");

            foreach (DataRow dr in dt5.Rows)
            {
                String NODO_LOCAL = dr["NODO_LOCAL"].ToString();
                String NODO_REMOTO = dr["NODO_REMOTO"].ToString();
                String RSS_LOCAL = dr["RSS_LOCAL"].ToString();
                String RSS_REMOTO = dr["RSS_REMOTO"].ToString();
                String TIEMPO_PROM = dr["TIEMPO_PROM"].ToString();
                String CAP_SUBIDA = dr["CAP_SUBIDA"].ToString();
                String CAP_BAJADA = dr["CAP_BAJADA"].ToString();
                String DISTANCIA_metros = dr["DISTANCIA_metros"].ToString();

                ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", NODO_LOCAL, 60, "B");
                ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", NODO_REMOTO, 60, "C");
                ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", RSS_LOCAL, 60, "D");
                ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", RSS_REMOTO, 60, "E");
                ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", "20", 60, "F");
                ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", "16 QAM", 60, "G");
                ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", TIEMPO_PROM, 60, "H");
                ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", "UL " + CAP_SUBIDA + "/" + "DL " + CAP_BAJADA, 60, "I");
                ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", DISTANCIA_metros, 60, "K");
                ExcelToolsBL.UpdateCell(excelGenerado, "11 DATOS GENERALES NODO A", FRECUENCIA_A + "MHz", 60, "L");

            }


            #endregion

            #region Datos Generales Nodo B

            ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", TIPO_NODO_B, 9, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", NODO_B, 12, "B");
            ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", NOMBRE_NODO_B, 12, "G");
            ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", NOMBRE_NODO_B, 17, "B");
            ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", UBIGEO_B, 17, "G");
            ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", DEPARTAMENTO_B, 20, "B");
            ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", PROVINCIA_B, 20, "E");
            ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", DISTRITO_B, 20, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", SERIE_PTP_450_B, 29, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", FRECUENCIA_B + " Mhz", 40, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", PIRE_EIRP_B + " dbm", 44, "I");

            if (ANTENA_MARCA_MODELO_B.Equals("INTEGRADA"))
            {
                ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", "Cambium Networks/5092HH", 46, "I");

            }
            else if (ANTENA_MARCA_MODELO_B.Equals("0.6"))
            {

                ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", "SHENGLU/SLU0652DD6B", 46, "I");
            }

            ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", GANANCIA_ANTENA_B, 47, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", ALTURA_ANTENA_B, 49, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", ELEVACION_B, 52, "I");
            ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", ALTITUD_msnm_B, 53, "I");

            foreach (DataRow dr in dt6.Rows)
            {
                String NODO_LOCAL = dr["NODO_LOCAL"].ToString();
                String NODO_REMOTO = dr["NODO_REMOTO"].ToString();
                String RSS_LOCAL = dr["RSS_LOCAL"].ToString();
                String RSS_REMOTO = dr["RSS_REMOTO"].ToString();
                String TIEMPO_PROM = dr["TIEMPO_PROM"].ToString();
                String CAP_SUBIDA = dr["CAP_SUBIDA"].ToString();
                String CAP_BAJADA = dr["CAP_BAJADA"].ToString();
                String DISTANCIA_metros = dr["DISTANCIA_metros"].ToString();

                ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", NODO_LOCAL, 60, "B");
                ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", NODO_REMOTO, 60, "C");
                ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", RSS_LOCAL, 60, "D");
                ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", RSS_REMOTO, 60, "E");
                ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", "20", 60, "F");
                ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", "16 QAM", 60, "G");
                ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", TIEMPO_PROM, 60, "H");
                ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", "UL " + CAP_SUBIDA + "/" + "DL " + CAP_BAJADA, 60, "I");
                ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", DISTANCIA_metros, 60, "K");
                ExcelToolsBL.UpdateCell(excelGenerado, "12 DATOS GENERALES NODO B", FRECUENCIA_B + "MHz", 60, "L");

            }


            #endregion


            #endregion

            #region Ruta del Zip
            String rutaNodo = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo;
            if (Directory.Exists(rutaNodo))
            {
                String rutaParcial = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\ACCESO\\PTP_HC-0XX1 (450i)\\";
                if (!Directory.Exists(rutaParcial)) Directory.CreateDirectory(rutaParcial);
                String rutaAlterna = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\ACCESO\\PTP_HC-0XX1 (450i)\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";
                File.Copy(excelGenerado, rutaAlterna, true);
            }
            #endregion

        }

        public void Anexo3ReporteFotograficoCMM4(String IdNodo, String IdTarea, String valorCadena1, String rutaPlantilla)
        {
            DBBaseDatos baseDatosDA = new DBBaseDatos();

            baseDatosDA.Configurar();
            baseDatosDA.Conectar();
            DataTable dt = new DataTable();
            try
            {
                baseDatosDA.CrearComando("USP_R_REPORTE_CMM4", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@CH_ID_TAREA", IdTarea, true);
                dt = baseDatosDA.EjecutarConsultaDataTable();

            }
            catch (Exception)
            {

                throw;
            }
            finally
            {

                baseDatosDA.Desconectar();
                baseDatosDA = null;
            }
            #region Valores

            byte[] FOTO_1_EQUIPO_GPS = (byte[])dt.Rows[0]["FOTO_1_EQUIPO_GPS"];
            MemoryStream mFOTO_1_EQUIPO_GPS = new MemoryStream(FOTO_1_EQUIPO_GPS);
            byte[] FOTO2_1_ATERRAMIENTO_GPS = (byte[])dt.Rows[0]["FOTO2_1_ATERRAMIENTO_GPS"];
            MemoryStream mFOTO2_1_ATERRAMIENTO_GPS = new MemoryStream(FOTO2_1_ATERRAMIENTO_GPS);
            byte[] FOTO2_2_ATERRAMIENTO_GPS = (byte[])dt.Rows[0]["FOTO2_2_ATERRAMIENTO_GPS"];
            MemoryStream mFOTO2_2_ATERRAMIENTO_GPS = new MemoryStream(FOTO2_2_ATERRAMIENTO_GPS);
            byte[] FOTO3_RECORRIDO_CABLE_CNT300 = (byte[])dt.Rows[0]["FOTO3_RECORRIDO_CABLE_CNT300"];
            MemoryStream mFOTO3_RECORRIDO_CABLE_CNT300 = new MemoryStream(FOTO3_RECORRIDO_CABLE_CNT300);
            byte[] FOTO4_1_ATERRAMIENTO_CABLE_CNT300 = (byte[])dt.Rows[0]["FOTO4_1_ATERRAMIENTO_CABLE_CNT300"];
            MemoryStream mFOTO4_1_ATERRAMIENTO_CABLE_CNT300 = new MemoryStream(FOTO4_1_ATERRAMIENTO_CABLE_CNT300);
            byte[] FOTO4_2_ATERRAMIENTO_CABLE_CNT300 = (byte[])dt.Rows[0]["FOTO4_2_ATERRAMIENTO_CABLE_CNT300"];
            MemoryStream mFOTO4_2_ATERRAMIENTO_CABLE_CNT300 = new MemoryStream(FOTO4_2_ATERRAMIENTO_CABLE_CNT300);
            byte[] FOTO5_1_ETIQUETADO_POE_CMM4 = (byte[])dt.Rows[0]["FOTO5_1_ETIQUETADO_POE_CMM4"];
            MemoryStream mFOTO5_1_ETIQUETADO_POE_CMM4 = new MemoryStream(FOTO5_1_ETIQUETADO_POE_CMM4);
            byte[] FOTO5_2_ETIQUETADO_POE_CMM4 = (byte[])dt.Rows[0]["FOTO5_2_ETIQUETADO_POE_CMM4"];
            MemoryStream mFOTO5_2_ETIQUETADO_POE_CMM4 = new MemoryStream(FOTO5_2_ETIQUETADO_POE_CMM4);
            byte[] FOTO6_1_PATCH_CORE = (byte[])dt.Rows[0]["FOTO6_1_PATCH_CORE"];
            MemoryStream mFOTO6_1_PATCH_CORE = new MemoryStream(FOTO6_1_PATCH_CORE);
            byte[] FOTO6_2_PATCH_CORE = (byte[])dt.Rows[0]["FOTO6_2_PATCH_CORE"];
            MemoryStream mFOTO6_2_PATCH_CORE = new MemoryStream(FOTO6_2_PATCH_CORE);
            byte[] FOTO7_1_POE_CMM4 = (byte[])dt.Rows[0]["FOTO7_1_POE_CMM4"];
            MemoryStream mFOTO7_1_POE_CMM4 = new MemoryStream(FOTO7_1_POE_CMM4);
            byte[] FOTO7_2_POE_CMM4 = (byte[])dt.Rows[0]["FOTO7_2_POE_CMM4"];
            MemoryStream mFOTO7_2_POE_CMM4 = new MemoryStream(FOTO7_2_POE_CMM4);
            byte[] FOTO8_TDK_LAMBDA = (byte[])dt.Rows[0]["FOTO8_TDK_LAMBDA"];
            MemoryStream mFOTO8_TDK_LAMBDA = new MemoryStream(FOTO8_TDK_LAMBDA);
            byte[] FOTO9_1_ENERGIA_TDK_LAMBDA = (byte[])dt.Rows[0]["FOTO9_1_ENERGIA_TDK_LAMBDA"];
            MemoryStream mFOTO9_1_ENERGIA_TDK_LAMBDA = new MemoryStream(FOTO9_1_ENERGIA_TDK_LAMBDA);
            byte[] FOTO9_2_ENERGIA_TDK_LAMBDA = (byte[])dt.Rows[0]["FOTO9_2_ENERGIA_TDK_LAMBDA"];
            MemoryStream mFOTO9_2_ENERGIA_TDK_LAMBDA = new MemoryStream(FOTO9_2_ENERGIA_TDK_LAMBDA);
            byte[] FOTO10_CONEXION_TDK_CMM4 = (byte[])dt.Rows[0]["FOTO10_CONEXION_TDK_CMM4"];
            MemoryStream mFOTO10_CONEXION_TDK_CMM4 = new MemoryStream(FOTO10_CONEXION_TDK_CMM4);




            #endregion

            // String usuarioWindows = Environment.UserName;
            String excelGenerado = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";
            File.Copy(rutaPlantilla, excelGenerado, true);

            #region Ingresando Valores

            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte Fotográfico", mFOTO_1_EQUIPO_GPS, "", 13, 3, 402, 340);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte Fotográfico", mFOTO2_1_ATERRAMIENTO_GPS, "", 13, 14, 373, 173);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte Fotográfico", mFOTO2_2_ATERRAMIENTO_GPS, "", 24, 14, 367, 161);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte Fotográfico", mFOTO3_RECORRIDO_CABLE_CNT300, "", 31, 4, 318, 341);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte Fotográfico", mFOTO4_1_ATERRAMIENTO_CABLE_CNT300, "", 31, 14, 372, 178);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte Fotográfico", mFOTO4_2_ATERRAMIENTO_CABLE_CNT300, "", 42, 14, 338, 147);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte Fotográfico", mFOTO5_1_ETIQUETADO_POE_CMM4, "", 49, 3, 338, 147);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte Fotográfico", mFOTO5_2_ETIQUETADO_POE_CMM4, "", 60, 4, 395, 166);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte Fotográfico", mFOTO6_1_PATCH_CORE, "", 49, 14, 405, 185);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte Fotográfico", mFOTO6_2_PATCH_CORE, "", 60, 14, 401, 157);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte Fotográfico", mFOTO7_1_POE_CMM4, "", 66, 5, 213, 181);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte Fotográfico", mFOTO7_2_POE_CMM4, "", 77, 4, 320, 164);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte Fotográfico", mFOTO8_TDK_LAMBDA, "", 66, 14, 397, 332);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte Fotográfico", mFOTO9_1_ENERGIA_TDK_LAMBDA, "", 82, 3, 399, 181);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte Fotográfico", mFOTO9_2_ENERGIA_TDK_LAMBDA, "", 93, 4, 299, 160);
            ExcelToolsBL.AddImageDocument(false, excelGenerado, "Reporte Fotográfico", mFOTO10_CONEXION_TDK_CMM4, "", 82, 15, 348, 341);

            #endregion

            #region Ruta del Zip
            String rutaNodo = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo;
            if (Directory.Exists(rutaNodo))
            {
                String rutaParcial = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\ACCESO\\PTP_HC-0XX1 (450i)\\";
                if (!Directory.Exists(rutaParcial)) Directory.CreateDirectory(rutaParcial);
                String rutaAlterna = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\" + IdNodo + "\\ACCESO\\PTP_HC-0XX1 (450i)\\" + IdNodo + " " + valorCadena1 + " " + IdTarea + ".xlsx";
                File.Copy(excelGenerado, rutaAlterna, true);
            }
            #endregion
        }
    }
}

