using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Data;
using DataAccess;
using BusinessEntity;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;

using System.IO.Compression;
using System.Web;


namespace BusinessLogic
{
   public  class ZipBL
    {

        DBBaseDatos baseDatosDA = new DBBaseDatos();

        public void DescargarZip(String IdNodo)
        {
            baseDatosDA.Configurar();
            baseDatosDA.Conectar();

            String ruta = "C:\\inetpub\\wwwroot\\SIAE_ARCHIVOS\\TEMPORAL\\";
            DataTable dt = new DataTable();
            DataTable dt1 = new DataTable();
            try
            {
                baseDatosDA.CrearComando("USP_R_ZIP_PRUEBA", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@VC_ID_NODO", IdNodo, true);
                dt = baseDatosDA.EjecutarConsultaDataTable();

                baseDatosDA.CrearComando("USP_R_EXCEL_IN_ZIP", CommandType.StoredProcedure);
                baseDatosDA.AsignarParametroCadena("@VC_ID_NODO", IdNodo, true);
                dt1 = baseDatosDA.EjecutarConsultaDataTable();


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

            String rutaExistente = ruta + IdNodo;
            if (Directory.Exists(rutaExistente))
            {

                Directory.Delete(rutaExistente, true);

            }

            if (File.Exists(ruta + IdNodo + ".zip"))
            {
                File.Delete(ruta + IdNodo + ".zip");
            }

            foreach (DataRow dr in dt.Rows)
            {
                String rutaCarpeta = dr["VC_RUTA_CARPETA"].ToString();
                String nombreArchivo = dr["VC_VALOR_CADENA1"].ToString();
                byte[] binario = (byte[])dr["VB_VALOR_BINARIO"];
                String extension = dr["VC_EXTENSION_ARCHIVO"].ToString();

                nombreArchivo = nombreArchivo.Substring(0,Math.Min(40,nombreArchivo.Length));
                nombreArchivo = nombreArchivo.Replace(":",""); //LA RUTA DE ARCHIVOS NO DEBE LLEVAR SIMBOLOS DE PUNTUACION
                nombreArchivo = nombreArchivo.Replace("/","");
                    
                String folder = IdNodo + rutaCarpeta;
                String rutaCompleta = Path.Combine(ruta, folder);


                Directory.CreateDirectory(rutaCompleta);
                File.WriteAllBytes(rutaCompleta + "\\" + nombreArchivo + extension, binario);             
            }

            #region  Codigo para descargar los excel
           String rutaPlantilla = "";
            ReporteDocumentosBL rd = new ReporteDocumentosBL();

            foreach (DataRow dr in dt1.Rows)
            {
                String IdDocumento = dr["CH_ID_DOCUMENTO"].ToString();
                String Tarea = dr["CH_ID_TAREA"].ToString();
                String NombreDocumento = dr["VC_VALOR_CADENA1"].ToString();
                String TipoNodoA = dr["CH_ID_TIP_NODO_A"].ToString();

                #region Valores para documentos

                switch (IdDocumento)
                {


                    case "000001":
                        rutaPlantilla = HttpContext.Current.Server.MapPath("~/Reportes/ActaInstalacionAceptacionProtocoloSectorial.xlsx");

                        rd.ActaInstalacionAceptacionProtocoloSectorial(IdNodo, Tarea, NombreDocumento, rutaPlantilla);

                        break;

                    case "000002":
                        rutaPlantilla = HttpContext.Current.Server.MapPath("~/Reportes/ActaInstalacionAceptacionProtocoloOmnidireccional.xlsx");

                        rd.ActaInstalacionAceptacionProtocoloOmnidireccional(IdNodo, Tarea, NombreDocumento, rutaPlantilla);

                        break;

                    case "000003":
                        rutaPlantilla = HttpContext.Current.Server.MapPath("~/Reportes/PruebaInterferencia.xlsx");

                        rd.PruebaInterferencia(IdNodo, Tarea, NombreDocumento, rutaPlantilla);

                        break;

                    case "000004":
                        rutaPlantilla = HttpContext.Current.Server.MapPath("~/Reportes/Anexo2InventarioPMP.xlsx");
                        rd.Anexo2InventarioPMP(IdNodo, Tarea, NombreDocumento, rutaPlantilla);

                        break;
                    case "000005":
                        rutaPlantilla = HttpContext.Current.Server.MapPath("~/Reportes/EstudioCampo.xlsx");
                        rd.EstudioDeCampo(IdNodo, Tarea, NombreDocumento, rutaPlantilla);

                        break;
                    case "000006":
                        rutaPlantilla = HttpContext.Current.Server.MapPath("~/Reportes/Anexo3ReporteFotográficoCMM4.xlsx");
                        rd.Anexo3ReporteFotograficoCMM4(IdNodo, Tarea, NombreDocumento, rutaPlantilla);

                        break;
                    case "000007":
                        rutaPlantilla = HttpContext.Current.Server.MapPath("~/Reportes/ProtocoloInstalacion.xlsx");
                        rd.ProtocoloInstalacion(IdNodo, Tarea, NombreDocumento, rutaPlantilla);

                        break;
                    case "000008":
                        rutaPlantilla = HttpContext.Current.Server.MapPath("~/Reportes/ActaInstalacionPTPLicenciado.xlsx");
                        rd.ActaInstalacionPTPLicenciado(IdNodo, Tarea, NombreDocumento, rutaPlantilla);

                        break;
                    case "000009":
                        rutaPlantilla = HttpContext.Current.Server.MapPath("~/Reportes/ActaInstalacionPTPNoLicenciado.xlsx");
                        rd.ActaInstalacionPTPNoLicenciado(IdNodo, Tarea, NombreDocumento, rutaPlantilla);

                        break;
                    case "000010":
                        rutaPlantilla = HttpContext.Current.Server.MapPath("~/Reportes/PruebasDeServicioDITGPTP.xlsx");
                        rd.PruebaServicioDITGPTP(IdNodo, Tarea, NombreDocumento, rutaPlantilla);

                        break;
                    case "000011":
                        rutaPlantilla = HttpContext.Current.Server.MapPath("~/Reportes/PruebasDeServicioDITGPMP.xlsx");
                        rd.PruebaServicioDITGPMP(IdNodo, Tarea, NombreDocumento, rutaPlantilla);

                        break;
                    case "000012":
                        rutaPlantilla = HttpContext.Current.Server.MapPath("~/Reportes/Anexo2InventarioPTP.xlsx");
                        rd.Anexo2InventarioPTP(IdNodo, Tarea, NombreDocumento, rutaPlantilla);

                        break;
                    case "000013":
                        if (TipoNodoA.Equals("000004") || TipoNodoA.Equals("000005"))
                        {
                            rutaPlantilla = HttpContext.Current.Server.MapPath("~/Reportes/ActaSeguridadAccesoDistritalDistribucion.xlsx");
                            rd.ActaSeguridadAcceso(IdNodo, Tarea, NombreDocumento, rutaPlantilla);
                        }
                        else
                        {
                            rutaPlantilla = HttpContext.Current.Server.MapPath("~/Reportes/ActaSeguridadAccesoIntermedioTerminal.xlsx");
                            rd.ActaSeguridadAcceso(IdNodo, Tarea, NombreDocumento, rutaPlantilla);
                        }

                        break;
                    case "000014":
                        if (TipoNodoA.Equals("000004") || TipoNodoA.Equals("000005"))
                        {
                            rutaPlantilla = HttpContext.Current.Server.MapPath("~/Reportes/ActaSeguridadDistribucionDistritalDistribucion.xlsx");
                            rd.ActaSeguridadDistribucion(IdNodo, Tarea, NombreDocumento, rutaPlantilla);
                        }
                        else
                        {
                            rutaPlantilla = HttpContext.Current.Server.MapPath("~/Reportes/ActaSeguridadDistribucionIntermedioTerminal.xlsx");
                            rd.ActaSeguridadDistribucion(IdNodo, Tarea, NombreDocumento, rutaPlantilla);
                        }

                        break;
                    case "000015":
                        rutaPlantilla = HttpContext.Current.Server.MapPath("~/Reportes/ActaInstalacionAceptacionProtocoloIIBB_A.xlsx");
                        rd.ActaInstalacionAceptacionProtocoloIIBB_A(IdNodo, Tarea, NombreDocumento, rutaPlantilla);

                        break;
                    case "000016":
                        rutaPlantilla = HttpContext.Current.Server.MapPath("~/Reportes/ActaInstalacionAceptacionProtocoloIIBB_B.xlsx");
                        rd.ActaInstalacionAceptacionProtocoloIIBB_B(IdNodo, Tarea, NombreDocumento, rutaPlantilla);

                        break;
                    case "000017":
                        rutaPlantilla = HttpContext.Current.Server.MapPath("~/Reportes/InstalaciondePozoaTierraTipoA.xlsx");
                        rd.InstalacionPozoTierraTipoA(IdNodo, Tarea, NombreDocumento, rutaPlantilla);

                        break;
                    case "000018":
                        rutaPlantilla = HttpContext.Current.Server.MapPath("~/Reportes/InstalaciondePozoaTierraTipoB.xlsx");
                        rd.InstalacionPozoTierraTipoB(IdNodo, Tarea, NombreDocumento, rutaPlantilla);

                        break;

                }

                #endregion
            }

            #endregion
            ZipFile.CreateFromDirectory(ruta + IdNodo,ruta + IdNodo + ".zip", CompressionLevel.Fastest, true);
        } 
    }

  
}
