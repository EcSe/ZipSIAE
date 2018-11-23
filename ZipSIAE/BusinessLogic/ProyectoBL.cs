using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BusinessEntity;
using DataAccess;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.IO;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;

namespace BusinessLogic
{
    public class ProyectoBL
    {
        public static void InsertarProyectoProceso(HiddenField hfArchivo,DropDownList ddlMetodo, 
            HtmlAnchor lnkLog,UsuarioBE UsuarioCreacion)
        {
            DBBaseDatos baseDatosDA = new DBBaseDatos();
            baseDatosDA.Configurar();
            baseDatosDA.Conectar();

            NodoBE Nodo = new NodoBE();
            InstitucionBeneficiariaBE InstitucionBeneficiaria = new InstitucionBeneficiariaBE();
            TareaBE Tarea = new TareaBE();
            PMPBE PMP = new PMPBE();
            IPPlanningPMPBE IPPlanningPMP = new IPPlanningPMPBE(); ;
            PTPBE PTP = new PTPBE();
            IPPlanningPTPBE IPPlanningPTP = new IPPlanningPTPBE();
            KitSIAEBE KitSIAE = new KitSIAEBE();
            DocumentoEquipamientoBE DocumentoEquipamiento = new DocumentoEquipamientoBE();
            DocumentoIPBE DocumentoIP = new DocumentoIPBE();
            DocumentoIPEquipamientoBE DocumentoIPEquipamiento = new DocumentoIPEquipamientoBE();
            DocumentoDetalleBE DocumentoDetalle = new DocumentoDetalleBE();

            EntidadDetalleBE conexionExcelBE = new EntidadDetalleBE();
            EntidadDetalleBE rutaTemporalBE = new EntidadDetalleBE();
            EntidadDetalleBE rutaVirtualTemporalBE = new EntidadDetalleBE();
            String strConexionExcel = "";
            OleDbConnection conexionExcel = new OleDbConnection();
            OleDbCommand Command;
            Int32 intFila = 0;
            Boolean blnErrorTabla = false;
            Boolean blnErrorCampo = false;
            Boolean blnErrorDato = false;
            Boolean blnErrorCampoTemp = false;
            Boolean blnErrorDatoTemp = false;

            rutaTemporalBE.Entidad.IdEntidad = "CONF";
            rutaTemporalBE.IdValor = "RUTA_TEMP";
            rutaTemporalBE = EntidadDetalleBL.ListarEntidadDetalle(rutaTemporalBE)[0];

            rutaVirtualTemporalBE.Entidad.IdEntidad = "CONF";
            rutaVirtualTemporalBE.IdValor = "RUTA_VIRT_TEMP";
            rutaVirtualTemporalBE = EntidadDetalleBL.ListarEntidadDetalle(rutaVirtualTemporalBE)[0];

            conexionExcelBE.Entidad.IdEntidad = "CONF";
            conexionExcelBE.IdValor = "CON_XLS";
            conexionExcelBE = EntidadDetalleBL.ListarEntidadDetalle(conexionExcelBE)[0];
            strConexionExcel = String.Format(conexionExcelBE.ValorCadena1, rutaTemporalBE.ValorCadena1 + "\\" + hfArchivo.Value);
            conexionExcel.ConnectionString = strConexionExcel;

            #region Para eliminar los registros del Estudio de Campo
            DocumentoDetalle.Documento = new DocumentoBE();
            DocumentoDetalle.Documento.Documento.IdValor = "000005";//ESTUDIO DE CAMPO
            #endregion

            conexionExcel.Open();
            //StreamWriter file = new StreamWriter(rutaTemporalBE.ValorCadena1 + "\\" + DateTime.Now.ToString("dd_MM_yyyy_HH_mm_ss") + "Log.txt");
            StreamWriter file = new StreamWriter(rutaTemporalBE.ValorCadena1 + "\\" + DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss") + "Log.txt");

            try
            {
                baseDatosDA.ComenzarTransaccion();

                #region Eliminamos toda la configuración
                NodoBL.EliminarFisicoNodoProceso(Nodo, baseDatosDA);
                InstitucionBeneficiariaBL.EliminarFisicoInstitucionBeneficiariaProceso(InstitucionBeneficiaria, baseDatosDA);
                TareaBL.EliminarFisicoTareaProceso(Tarea, baseDatosDA);
                PMPBL.EliminarFisicoPMPProceso(PMP, baseDatosDA);
                IPPlanningPMPBL.EliminarFisicoIPPlanningPMPProceso(IPPlanningPMP, baseDatosDA);
                PTPBL.EliminarFisicoPTPProceso(PTP, baseDatosDA);
                IPPlanningPTPBL.EliminarFisicoIPPlanningPTPProceso(IPPlanningPTP, baseDatosDA);
                KitSIAEBL.EliminarFisicoKitSIAEProceso(KitSIAE, baseDatosDA);
                DocumentoEquipamientoBL.EliminarFisicoDocumentoEquipamientoProceso(DocumentoEquipamiento, baseDatosDA);
                DocumentoIPBL.EliminarFisicoDocumentoIPProceso(DocumentoIP, baseDatosDA);
                DocumentoIPEquipamientoBL.EliminarFisicoDocumentoIPEquipamientoProceso(DocumentoIPEquipamiento, baseDatosDA);
                DocumentoDetalleBL.EliminarFisicoEntidadDetalleProceso(DocumentoDetalle, baseDatosDA);
                #endregion

                #region Insertamos toda la configuracion

                #region Validamos la hoja DATOS GENERALES NODO
                intFila = 2;
                file.WriteLine("DATOS GENERALES NODO");
                file.WriteLine("--------------------");
                Command = new OleDbCommand("SELECT * FROM [DATOS GENERALES NODO$]", conexionExcel);
                try
                {
                    blnErrorTabla = false;
                    DbDataReader reader = Command.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            Object objValor = null;
                            blnErrorDato = false;
                            blnErrorCampo = false;
                            blnErrorDatoTemp = false;
                            blnErrorCampoTemp = false;

                            Nodo = new NodoBE();

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "CODIGO SITIO", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                Nodo.IdNodo = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "REGION", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                Nodo.Region.Nombre = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "NOMBRE DEL NODO", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                Nodo.Nombre = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "REGION DEPARTAMENTO", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                Nodo.Localidad.Distrito.Provincia.Departamento.Nombre = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "PROVINCIA", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                Nodo.Localidad.Distrito.Provincia.Nombre = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "DISTRITO", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                Nodo.Localidad.Distrito.Nombre = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "LOCALIDAD", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                Nodo.Localidad.Nombre = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "LATITUD", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                Nodo.Latitud = Convert.ToDouble(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "LONGITUD", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                Nodo.Longitud = Convert.ToDouble(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "ANILLO", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                Nodo.Anillo = Convert.ToInt32(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "TORRE ALTURA (m)", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                Nodo.AlturaTorre = Convert.ToInt32(objValor);

                            #region Si no hay errores de campo o de dato intentamos procesar la fila.
                            if (!blnErrorCampo && !blnErrorDato)
                            {
                                try
                                {
                                    if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                    {
                                        Nodo.UsuarioCreacion = UsuarioCreacion;
                                        NodoBL.InsertarNodoProceso(Nodo, baseDatosDA);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    blnErrorDato = true;
                                    file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                }
                            }
                            #endregion
                            intFila++;
                        }
                    }
                    else
                    {
                        file.WriteLine("La tabla DATOS_GENERALES_NODO no tiene registros.");
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    if (ex.GetType().FullName.Equals("System.Data.OleDb.OleDbException") && ((System.Data.OleDb.OleDbException)ex).ErrorCode.Equals(-2147467259))
                    {
                        file.WriteLine("No existe la tabla DATOS_GENERALES_NODO");
                    }
                    else
                    {
                        file.WriteLine(ex.Message);
                    }
                    blnErrorTabla = true;
                }
                #endregion

                #region Validamos la hoja PMP sólo IIBB
                intFila = 2;
                file.WriteLine("");
                file.WriteLine("");
                file.WriteLine("PMP - SÓLO IIBB");
                file.WriteLine("---------------");
                Command = new OleDbCommand("SELECT * FROM [PMP$]", conexionExcel);
                try
                {
                    blnErrorTabla = false;
                    DbDataReader reader = Command.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            Object objValor = null;
                            blnErrorDato = false;
                            blnErrorCampo = false;
                            blnErrorDatoTemp = false;
                            blnErrorCampoTemp = false;

                            #region Validamos los campos de la InstitucionBeneficiaria
                            //InstitucionBeneficiariaBE
                            InstitucionBeneficiaria = new InstitucionBeneficiariaBE();
                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "CODIGO IIBB (SITE B)", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                InstitucionBeneficiaria.IdInstitucionBeneficiaria = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "INSTITUCION BENEFICIARIA (SITE B)", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                InstitucionBeneficiaria.Nombre = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "LATITUD (SITE B)", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                InstitucionBeneficiaria.Latitud = Convert.ToDouble(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "LONGITUD (SITE B)", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                InstitucionBeneficiaria.Longitud = Convert.ToDouble(objValor);
                            #endregion

                            #region Si no hay errores de campo o de dato intentamos procesar la fila.
                            if (!blnErrorCampo && !blnErrorDato)
                            {
                                try
                                {

                                    if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                    {
                                        InstitucionBeneficiaria.UsuarioCreacion = UsuarioCreacion;
                                        InstitucionBeneficiariaBL.InsertarInstitucionBeneficiariaProceso(InstitucionBeneficiaria, baseDatosDA);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    blnErrorDato = true;
                                    file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                }
                            }
                            #endregion
                            intFila++;
                        }
                    }
                    else
                    {
                        file.WriteLine("La tabla PMP no tiene registros.");
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    if (ex.GetType().FullName.Equals("System.Data.OleDb.OleDbException") && ((System.Data.OleDb.OleDbException)ex).ErrorCode.Equals(-2147467259))
                    {
                        file.WriteLine("No existe la tabla PMP");
                    }
                    else
                    {
                        file.WriteLine(ex.Message);
                    }
                    blnErrorTabla = true;
                }
                #endregion

                #region Validamos la hoja DOCUMENTACION
                intFila = 2;
                file.WriteLine("");
                file.WriteLine("");
                file.WriteLine("DOCUMENTACION");
                file.WriteLine("-------------");
                Command = new OleDbCommand("SELECT * FROM [DOCUMENTACION$]", conexionExcel);
                try
                {
                    blnErrorTabla = false;
                    DbDataReader reader = Command.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            Object objValor = null;
                            blnErrorDato = false;
                            blnErrorCampo = false;
                            blnErrorDatoTemp = false;
                            blnErrorCampoTemp = false;

                            #region Validamos los campos de Tarea
                            Tarea = new TareaBE();
                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "Link_ID", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                Tarea.IdTarea = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "IsoNo", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                Tarea.IdIsoNodo = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "LINK TYPE", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                Tarea.TipoTarea.ValorCadena1 = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, true, "SUBCONTRATA", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp && !objValor.Equals(DBNull.Value))
                                Tarea.Contratista.ValorCadena2 = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<DateTime>(reader, intFila, true, "Inicio Instalacion", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp && !objValor.Equals(DBNull.Value))
                                Tarea.InicioInstalacion = Convert.ToDateTime(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<DateTime>(reader, intFila, true, "Fin Instalacion", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp && !objValor.Equals(DBNull.Value))
                                Tarea.FinInstalacion = Convert.ToDateTime(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "PROYECTO", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                Tarea.Proyecto.ValorCadena1 = Convert.ToString(objValor);

                            if (Tarea.TipoTarea.ValorCadena1.Equals("CPE-A") || Tarea.TipoTarea.ValorCadena1.Equals("CPE-B"))
                            {
                                Tarea.TipoNodoA.ValorCadena1 = null;
                            }
                            else
                            {
                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "Tipo de nodo A", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp && !objValor.Equals(DBNull.Value))
                                    Tarea.TipoNodoA.ValorCadena1 = Convert.ToString(objValor);
                            }

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "Codigo A", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                Tarea.NodoIIBBA.IdNodo = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, true, "Tipo de nodo B", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp && !objValor.Equals(DBNull.Value))
                                Tarea.TipoNodoB.ValorCadena1 = Convert.ToString(objValor);

                            if (Tarea.TipoTarea.ValorCadena1.Equals("CPE-A") || Tarea.TipoTarea.ValorCadena1.Equals("CPE-B"))
                            {
                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "Codigo B", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                    Tarea.NodoB.IdNodo = Convert.ToString(objValor);
                            }
                            else
                            {
                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, true, "Codigo B", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp && !objValor.Equals(DBNull.Value))
                                    Tarea.NodoB.IdNodo = Convert.ToString(objValor);
                            }

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, true, "SECTOR AP", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp && !objValor.Equals(DBNull.Value))
                                Tarea.IdSectorAP = Convert.ToString(objValor);

                            #endregion

                            #region Si no hay errores de campo o de dato intentamos procesar la fila.
                            if (!blnErrorCampo && !blnErrorDato)
                            {
                                try
                                {
                                    if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                    {
                                        Tarea.UsuarioCreacion = UsuarioCreacion;
                                        TareaBL.InsertarTareaProceso(Tarea, baseDatosDA);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    blnErrorDato = true;
                                    file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                }
                            }
                            #endregion
                            intFila++;
                        }
                    }
                    else
                    {
                        file.WriteLine("La tabla DOCUMENTACION no tiene registros.");
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    if (ex.GetType().FullName.Equals("System.Data.OleDb.OleDbException") && ((System.Data.OleDb.OleDbException)ex).ErrorCode.Equals(-2147467259))
                    {
                        file.WriteLine("No existe la tabla DOCUMENTACION");
                    }
                    else
                    {
                        file.WriteLine(ex.Message);
                    }
                    blnErrorTabla = true;
                }
                #endregion

                #region Validamos la hoja PMP
                intFila = 2;
                file.WriteLine("");
                file.WriteLine("");
                file.WriteLine("PMP");
                file.WriteLine("---");
                Command = new OleDbCommand("SELECT * FROM [PMP$]", conexionExcel);
                try
                {
                    blnErrorTabla = false;
                    DbDataReader reader = Command.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            Object objValor = null;
                            blnErrorDato = false;
                            blnErrorCampo = false;
                            blnErrorDatoTemp = false;
                            blnErrorCampoTemp = false;

                            #region Validamos los campos de la InstitucionBeneficiaria
                            InstitucionBeneficiaria = new InstitucionBeneficiariaBE();
                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "CODIGO IIBB (SITE B)", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                InstitucionBeneficiaria.IdInstitucionBeneficiaria = Convert.ToString(objValor);

                            #endregion

                            #region Validamos los campos de PMP
                            PMP = new PMPBE();
                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "CODIGO AP (SITE A)", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PMP.Nodo.IdNodo = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "Antenna Model (SITE A)", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PMP.ModeloAntenaNodo.ValorCadena1 = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "Antenna Gain (SITE A)", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PMP.GananciaAntenaNodo = Convert.ToInt32(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "Antenna Height (SITE A)", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PMP.AlturaAntenaNodo = Convert.ToInt32(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "Azimuth (SITE A) (°)", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PMP.AzimuthAntenaNodo = Convert.ToInt32(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "Elevation (SITE A) (°)", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PMP.ElevacionAntenaNodo = Convert.ToInt32(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "EIRP (SITE A) (dBm)", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PMP.EIRPAntenaNodo = Convert.ToInt32(objValor);

                            #endregion

                            #region Validamos los campos de PMPDetalle
                            PMPDetalleBE PMPDetalle = new PMPDetalleBE();

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "CODIGO IIBB (SITE B)", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PMPDetalle.InstitucionBeneficiaria.IdInstitucionBeneficiaria = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "SECTOR DEL AP", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PMPDetalle.SectorIIBB = Convert.ToInt32(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "Azimuth (SITE B) (°)", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PMPDetalle.AzimuthAntenaIIBB = Convert.ToDouble(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "Elevation (SITE B) (°)", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PMPDetalle.ElevacionAntenaIIBB = Convert.ToDouble(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "Tx Power (SITE B) (dBm)", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PMPDetalle.TXTorreIIBB = Convert.ToInt32(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "EIRP (SITE B) (dBm)", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PMPDetalle.EIRPAntenaIIBB = Convert.ToDouble(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "Rx Level (SITE A) (dBm)", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PMPDetalle.NivelRXNodo = Convert.ToDouble(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "Rx Level (SITE B) (dBm)", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PMPDetalle.NivelRXIIBB = Convert.ToDouble(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "Fade Margin (SITE A) (dB)", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PMPDetalle.FadeMarginNodo = Convert.ToDouble(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "Fade Margin (SITE B) (dB)", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PMPDetalle.FadeMarginIIBB = Convert.ToDouble(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "Availability (%)", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PMPDetalle.Disponibilidad = Convert.ToDouble(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "Distance (Km)", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PMPDetalle.Distancia = Convert.ToDouble(objValor);

                            #endregion

                            #region Si no hay errores de campo o de dato intentamos procesar la fila.
                            if (!blnErrorCampo && !blnErrorDato)
                            {
                                try
                                {
                                    PMP.Detalles.Add(PMPDetalle);
                                    if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                    {
                                        PMP.UsuarioCreacion = UsuarioCreacion;
                                        PMPBL.InsertarPMPProceso(PMP, baseDatosDA);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    blnErrorDato = true;
                                    file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                }
                            }
                            #endregion
                            intFila++;
                        }
                    }
                    else
                    {
                        file.WriteLine("La tabla PMP no tiene registros.");
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    if (ex.GetType().FullName.Equals("System.Data.OleDb.OleDbException") && ((System.Data.OleDb.OleDbException)ex).ErrorCode.Equals(-2147467259))
                    {
                        file.WriteLine("No existe la tabla PMP");
                    }
                    else
                    {
                        file.WriteLine(ex.Message);
                    }
                    blnErrorTabla = true;
                }
                #endregion

                #region Validamos la hoja IP PLANNING PMP
                intFila = 2;
                file.WriteLine("");
                file.WriteLine("");
                file.WriteLine("IP PLANNING PMP");
                file.WriteLine("---------------");
                Command = new OleDbCommand("SELECT * FROM [IP PLANNING PMP$]", conexionExcel);
                try
                {
                    blnErrorTabla = false;
                    DbDataReader reader = Command.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            Object objValor = null;
                            blnErrorDato = false;
                            blnErrorCampo = false;
                            blnErrorDatoTemp = false;
                            blnErrorCampoTemp = false;

                            #region Validamos los campos de IPPlanningPMP
                            IPPlanningPMP = new IPPlanningPMPBE();
                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "ESTACION A", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                IPPlanningPMP.Nodo.IdNodo = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "IP A", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                IPPlanningPMP.IPNodo = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "DEFAULT GATEWAY A Y B", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                IPPlanningPMP.DefaultGateway = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "IP Conexión Local", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                IPPlanningPMP.IPConexionLocal = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "PUERTO A", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                IPPlanningPMP.PuertoNodo = Convert.ToString(objValor);

                            #endregion

                            #region Validamos los campos de IPPlanningPMPDetalle
                            IPPlanningPMPDetalleBE IPPlanningPMPDetalle = new IPPlanningPMPDetalleBE();
                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "ESTACION B", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                IPPlanningPMPDetalle.InstitucionBeneficiaria.IdInstitucionBeneficiaria = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "IP B", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                IPPlanningPMPDetalle.IPIIBB = Convert.ToString(objValor);

                            #endregion

                            #region Si no hay errores de campo o de dato intentamos procesar la fila.
                            if (!blnErrorCampo && !blnErrorDato)
                            {
                                try
                                {
                                    //IPPlanningPMPDetalle.IPPlanningPMP = IPPlanningPMP;
                                    IPPlanningPMP.Detalles.Add(IPPlanningPMPDetalle);
                                    if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                    {
                                        IPPlanningPMP.UsuarioCreacion = UsuarioCreacion;
                                        IPPlanningPMPBL.InsertarIPPlanningPMPProceso(IPPlanningPMP, baseDatosDA);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    blnErrorDato = true;
                                    file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                }
                            }
                            #endregion
                            intFila++;
                        }
                    }
                    else
                    {
                        file.WriteLine("La tabla PMP no tiene registros.");
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    if (ex.GetType().FullName.Equals("System.Data.OleDb.OleDbException") && ((System.Data.OleDb.OleDbException)ex).ErrorCode.Equals(-2147467259))
                    {
                        file.WriteLine("No existe la tabla IP PLANNING PMP");
                    }
                    else
                    {
                        file.WriteLine(ex.Message);
                    }
                    blnErrorTabla = true;
                }
                #endregion

                #region Validamos la hoja PTP
                intFila = 2;
                file.WriteLine("");
                file.WriteLine("");
                file.WriteLine("PTP");
                file.WriteLine("---");
                Command = new OleDbCommand("SELECT * FROM [PTP$]", conexionExcel);
                try
                {
                    blnErrorTabla = false;
                    DbDataReader reader = Command.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            Object objValor = null;
                            blnErrorDato = false;
                            blnErrorCampo = false;
                            blnErrorDatoTemp = false;
                            blnErrorCampoTemp = false;

                            #region Validamos los campos de PTP
                            PTP = new PTPBE();
                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "Call sign S1", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PTP.NodoA.IdNodo = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "COTA (msnm) S1", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PTP.CotaNodoA = Convert.ToDouble(objValor);

                            #endregion

                            #region Validamos los campos de PTPDetalle
                            PTPDetalleBE PTPDetalle = new PTPDetalleBE();
                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "Call sign S2", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PTPDetalle.NodoB.IdNodo = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "True azimuth (°) S1", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PTPDetalle.AzimuthNodoA = Convert.ToDouble(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "True azimuth (°) S2", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PTPDetalle.AzimuthNodoB = Convert.ToDouble(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "ELEVACIÓN (°) S1", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PTPDetalle.ElevacionNodoA = Convert.ToDouble(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "ELEVACIÓN (°) S2", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PTPDetalle.ElevacionNodoB = Convert.ToDouble(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "Distancia (km)", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PTPDetalle.Distancia = Convert.ToDouble(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "COTA (msnm) S2", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PTPDetalle.CotaNodoB = Convert.ToDouble(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "TR Antenna model S1", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PTPDetalle.ModeloAntenaNodoA = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "TR Antenna model S2", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PTPDetalle.ModeloAntenaNodoB = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "TR Antenna diameter (m) S1", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PTPDetalle.DiametroAntenaNodoA = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "TR Antenna diameter (m) S2", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PTPDetalle.DiametroAntenaNodoB = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "TR Antenna height (m) S1", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PTPDetalle.AlturaAntenaNodoA = Convert.ToInt32(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "TR Antenna height (m) S2", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PTPDetalle.AlturaAntenaNodoB = Convert.ToInt32(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "TR Antenna gain (dBi) S1", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PTPDetalle.GananciaAntenaNodoA = Convert.ToDouble(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "TR Antenna gain (dBi) S2", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PTPDetalle.GananciaAntenaNodoB = Convert.ToDouble(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "#1 Channel ID S1", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PTPDetalle.IdCanalNodoA = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "#1 Channel ID S2", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PTPDetalle.IdCanalNodoB = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "#1 Design frequency S1", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PTPDetalle.DisenoFrecuenciaNodoA = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "#1 Design frequency S2", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PTPDetalle.DisenoFrecuenciaNodoB = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "Polarization", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PTPDetalle.Polarizacion.ValorCadena1 = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "Radio model S1", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PTPDetalle.ModeloRadioNodoA = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "Emission designator S1", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PTPDetalle.DesignadorEmisionNodoA = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "TX power (dBm) S1", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PTPDetalle.PotenciaTorreNodoA = Convert.ToInt32(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "TX power (dBm) S2", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PTPDetalle.PotenciaTorreNodoB = Convert.ToInt32(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "EIRP (dBm) S1", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PTPDetalle.EIRPNodoA = Convert.ToDouble(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "EIRP (dBm) S2", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PTPDetalle.EIRPNodoB = Convert.ToDouble(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "RX threshold level (dBm) S1", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PTPDetalle.NivelUmbralNodoA = Convert.ToDouble(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "Receive signal (dBm) S1", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PTPDetalle.SenalRecepcionNodoA = Convert.ToDouble(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "Receive signal (dBm) S2", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PTPDetalle.SenalRecepcionNodoB = Convert.ToDouble(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "Effective fade margin (dB) S1", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PTPDetalle.MargenEfectividadDesvanecimientoNodoA = Convert.ToDouble(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "Effective fade margin (dB) S2", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PTPDetalle.MargenEfectividadDesvanecimientoNodoB = Convert.ToDouble(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "Annual multipath availability (%) S1", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PTPDetalle.DisponibilidadAnualMultirutasNodoA = Convert.ToDouble(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "Annual rain availability (%) S1", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PTPDetalle.DisponibilidadAnualLluviaNodoA = Convert.ToDouble(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, false, "Annual rain + multipath availability (%)", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                PTPDetalle.DisponibilidadAnualMultirutasLluvia = Convert.ToDouble(objValor);

                            #endregion

                            #region Si no hay errores de campo o de dato intentamos procesar la fila.
                            if (!blnErrorCampo && !blnErrorDato)
                            {
                                try
                                {
                                    //PTPDetalle.PTP = PTP;
                                    PTP.Detalles.Add(PTPDetalle);
                                    if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                    {
                                        PTP.UsuarioCreacion = UsuarioCreacion;
                                        PTPBL.InsertarPTPProceso(PTP, baseDatosDA);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    blnErrorDato = true;
                                    file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                }
                            }
                            #endregion

                            intFila++;
                        }
                    }
                    else
                    {
                        file.WriteLine("La tabla PMP no tiene registros.");
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    if (ex.GetType().FullName.Equals("System.Data.OleDb.OleDbException") && ((System.Data.OleDb.OleDbException)ex).ErrorCode.Equals(-2147467259))
                    {
                        file.WriteLine("No existe la tabla PTP");
                    }
                    else
                    {
                        file.WriteLine(ex.Message);
                    }
                    blnErrorTabla = true;
                }
                #endregion

                #region Validamos la hoja IP PLANNING PTP
                intFila = 2;
                file.WriteLine("");
                file.WriteLine("");
                file.WriteLine("IP PLANNING PTP");
                file.WriteLine("---------------");
                Command = new OleDbCommand("SELECT * FROM [IP PLANNING PTP$]", conexionExcel);
                try
                {
                    blnErrorTabla = false;
                    DbDataReader reader = Command.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            Object objValor = null;
                            blnErrorDato = false;
                            blnErrorCampo = false;
                            blnErrorDatoTemp = false;
                            blnErrorCampoTemp = false;

                            #region Validamos los campos de IPPlanningPTP
                            IPPlanningPTP = new IPPlanningPTPBE();
                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "ESTACION A", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                IPPlanningPTP.NodoA.IdNodo = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "DEFAULT GATEWAY A Y B", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                IPPlanningPTP.DefaultGateway = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "IP Conexión Local", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                IPPlanningPTP.IPConexionLocal = Convert.ToString(objValor);

                            #endregion

                            #region Validamos los campos de IPPlanningPTPDetalle
                            IPPlanningPTPDetalleBE IPPlanningPTPDetalle = new IPPlanningPTPDetalleBE();
                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, true, "ESTACION B", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp && !objValor.Equals(DBNull.Value))
                                IPPlanningPTPDetalle.NodoB.IdNodo = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "IP A", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                IPPlanningPTPDetalle.IPNodoA = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, true, "IP B", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp && !objValor.Equals(DBNull.Value))
                                IPPlanningPTPDetalle.IPNodoB = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "PUERTO A", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                IPPlanningPTPDetalle.PuertoNodoA = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, true, "PUERTO B", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp && !objValor.Equals(DBNull.Value))
                                IPPlanningPTPDetalle.PuertoNodoB = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<Double>(reader, intFila, true, "COLOR CODE", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp && !objValor.Equals(DBNull.Value))
                                IPPlanningPTPDetalle.CodigoColor = Convert.ToInt32(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, true, "MAESTRO", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp && !objValor.Equals(DBNull.Value))
                                IPPlanningPTPDetalle.NodoMaestro.IdNodo = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, true, "SINCRONISMO", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp && !objValor.Equals(DBNull.Value))
                                IPPlanningPTPDetalle.Sincronismo.ValorCadena1 = Convert.ToString(objValor);


                            #endregion

                            #region Si no hay errores de campo o de dato intentamos procesar la fila.
                            if (!blnErrorCampo && !blnErrorDato)
                            {
                                try
                                {
                                    //IPPlanningPTPDetalle.IPPlanningPTP = IPPlanningPTP;
                                    IPPlanningPTP.Detalles.Add(IPPlanningPTPDetalle);
                                    if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                    {
                                        IPPlanningPTP.UsuarioCreacion = UsuarioCreacion;
                                        IPPlanningPTPBL.InsertarIPPlanningPTPProceso(IPPlanningPTP, baseDatosDA);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    blnErrorDato = true;
                                    file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                }
                            }
                            #endregion
                            intFila++;
                        }
                    }
                    else
                    {
                        file.WriteLine("La tabla PMP no tiene registros.");
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    if (ex.GetType().FullName.Equals("System.Data.OleDb.OleDbException") && ((System.Data.OleDb.OleDbException)ex).ErrorCode.Equals(-2147467259))
                    {
                        file.WriteLine("No existe la tabla IP PLANNING PTP");
                    }
                    else
                    {
                        file.WriteLine(ex.Message);
                    }
                    blnErrorTabla = true;
                }
                #endregion

                #region Validamos la hoja KIT SIAE
                intFila = 2;
                file.WriteLine("");
                file.WriteLine("");
                file.WriteLine("KIT SIAE");
                file.WriteLine("--------");
                Command = new OleDbCommand("SELECT * FROM [KIT SIAE$] WHERE TIPO_MOV = 'OUT'", conexionExcel);
                try
                {
                    blnErrorTabla = false;
                    DbDataReader reader = Command.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            Object objValor = null;
                            blnErrorDato = false;
                            blnErrorCampo = false;
                            blnErrorDatoTemp = false;
                            blnErrorCampoTemp = false;

                            #region Validamos los campos de KitSIAE
                            KitSIAE = new KitSIAEBE();
                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "S/N", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                            {
                                KitSIAE.SerieKit = Convert.ToString(objValor);
                                KitSIAE.SerieKit = KitSIAE.SerieKit.Substring(KitSIAE.SerieKit.Length - 4);
                            }

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "CODIGO_GILAT", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                KitSIAE.CodigoGilat = Convert.ToString(objValor);


                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "LINK ID", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                KitSIAE.Tarea.IdTarea = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "SITE CODE", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                KitSIAE.Tarea.NodoIIBBA.IdNodo = Convert.ToString(objValor);

                            #endregion

                            #region Si no hay errores de campo o de dato intentamos procesar la fila.
                            if (!blnErrorCampo && !blnErrorDato)
                            {
                                try
                                {
                                    if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                    {
                                        KitSIAE.UsuarioCreacion = UsuarioCreacion;
                                        KitSIAEBL.InsertarKitSIAEProceso(KitSIAE, baseDatosDA);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    blnErrorDato = true;
                                    file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                }
                            }
                            #endregion
                            intFila++;
                        }
                    }
                    else
                    {
                        file.WriteLine("La tabla KIT SIAE no tiene registros.");
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    if (ex.GetType().FullName.Equals("System.Data.OleDb.OleDbException") && ((System.Data.OleDb.OleDbException)ex).ErrorCode.Equals(-2147467259))
                    {
                        file.WriteLine("No existe la tabla KIT SIAE");
                    }
                    else
                    {
                        file.WriteLine(ex.Message);
                    }
                    blnErrorTabla = true;
                }
                #endregion

                #region Validamos la hoja EQUIPAMIENTO SIAE
                intFila = 2;
                file.WriteLine("");
                file.WriteLine("");
                file.WriteLine("EQUIPAMIENTO SIAE");
                file.WriteLine("-----------------");
                Command = new OleDbCommand("SELECT * FROM [EQUIPAMIENTO SIAE$] WHERE TIPO_MOV = 'OUT'", conexionExcel);
                try
                {
                    blnErrorTabla = false;
                    DbDataReader reader = Command.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            Object objValor = null;
                            blnErrorDato = false;
                            blnErrorCampo = false;
                            blnErrorDatoTemp = false;
                            blnErrorCampoTemp = false;

                            #region Validamos los campos de DocumentoEquipamiento
                            DocumentoEquipamiento = new DocumentoEquipamientoBE();
                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "CODIGO_SIAE", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoEquipamiento.Equipamiento.IdValor = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "S/N", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoEquipamiento.SerieEquipamiento = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "LINK ID", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoEquipamiento.Documento.Tarea.IdTarea = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "SITE CODE", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoEquipamiento.Documento.Tarea.NodoIIBBA.IdNodo = Convert.ToString(objValor);

                            DocumentoEquipamiento.IdEmpresa = "S";//SIAE

                            //Obtenemos el tipo de tarea
                            Tarea = new TareaBE();
                            Tarea.IdTarea = DocumentoEquipamiento.Documento.Tarea.IdTarea;
                            try
                            {
                                Tarea = TareaBL.ListarTareas(Tarea,"S", baseDatosDA)[0];
                            }
                            catch (Exception ex)
                            {
                                Tarea.NodoB.IdNodo = "";
                            }

                            #endregion

                            #region Si no hay errores de campo o de dato intentamos procesar la fila.
                            if (!(Tarea.TipoTarea.IdValor.Equals("000012") && DocumentoEquipamiento.Equipamiento.IdValor.Equals("D60078")))
                            {
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoEquipamiento.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoEquipamientoBL.InsertarDocumentoEquipamientoProceso(DocumentoEquipamiento, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                            }
                            #endregion

                            intFila++;
                        }
                    }
                    else
                    {
                        file.WriteLine("La tabla EQUIPAMIENTO SIAE no tiene registros.");
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    if (ex.GetType().FullName.Equals("System.Data.OleDb.OleDbException") && ((System.Data.OleDb.OleDbException)ex).ErrorCode.Equals(-2147467259))
                    {
                        file.WriteLine("No existe la tabla EQUIPAMIENTO SIAE");
                    }
                    else
                    {
                        file.WriteLine(ex.Message);
                    }
                    blnErrorTabla = true;
                }
                #endregion

                #region Validamos la hoja EQUIPAMIENTO SIAE ALIMENTACION
                intFila = 2;
                file.WriteLine("");
                file.WriteLine("");
                file.WriteLine("EQUIPAMIENTO SIAE ALIMENTACION");
                file.WriteLine("------------------------------");
                Command = new OleDbCommand("SELECT * FROM [EQUIPAMIENTO SIAE$] WHERE TIPO_MOV = 'OUT' AND CODIGO_SIAE IN ('D60078','D60077','ICA0071') ", conexionExcel);
                try
                {
                    blnErrorTabla = false;
                    DbDataReader reader = Command.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            Object objValor = null;
                            blnErrorDato = false;
                            blnErrorCampo = false;
                            blnErrorDatoTemp = false;
                            blnErrorCampoTemp = false;

                            #region Validamos los campos de DocumentoEquipamiento
                            DocumentoEquipamiento = new DocumentoEquipamientoBE();
                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "CODIGO_SIAE", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoEquipamiento.Equipamiento.IdValor = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "S/N", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoEquipamiento.SerieEquipamiento = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "LINK ID", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoEquipamiento.Documento.Tarea.IdTarea = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "SITE CODE", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoEquipamiento.Documento.Tarea.NodoIIBBA.IdNodo = Convert.ToString(objValor);

                            DocumentoEquipamiento.IdEmpresa = "S";//SIAE

                            //Obtenemos el tipo de tarea
                            Tarea = new TareaBE();
                            Tarea.IdTarea = DocumentoEquipamiento.Documento.Tarea.IdTarea;
                            try
                            {
                                Tarea = TareaBL.ListarTareas(Tarea, "S", baseDatosDA)[0];
                            }
                            catch (Exception ex)
                            {
                                Tarea.NodoB.IdNodo = "";
                            }

                            #endregion

                            #region Si no hay errores de campo o de dato intentamos procesar la fila.
                            if (Tarea.TipoTarea.IdValor.Equals("000012"))
                            {
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoEquipamiento.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoEquipamientoBL.InsertarDocumentoEquipamientoAlimentacionProceso(DocumentoEquipamiento, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                            }
                            #endregion

                            intFila++;
                        }
                    }
                    else
                    {
                        file.WriteLine("La tabla EQUIPAMIENTO SIAE no tiene registros.");
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    if (ex.GetType().FullName.Equals("System.Data.OleDb.OleDbException") && ((System.Data.OleDb.OleDbException)ex).ErrorCode.Equals(-2147467259))
                    {
                        file.WriteLine("No existe la tabla EQUIPAMIENTO SIAE");
                    }
                    else
                    {
                        file.WriteLine(ex.Message);
                    }
                    blnErrorTabla = true;
                }
                #endregion

                #region Validamos la hoja EQUIPAMIENTO AIO
                intFila = 2;
                file.WriteLine("");
                file.WriteLine("");
                file.WriteLine("EQUIPAMIENTO AIO");
                file.WriteLine("----------------");
                Command = new OleDbCommand("SELECT * FROM [EQUIPAMIENTO AIO$] WHERE [SERIE DE KIT] <> 'REPUESTOS'", conexionExcel);
                //AND [NUMERO DE SERIE] <> '-' AND [NUMERO DE SERIE] <> ''
                try
                {
                    blnErrorTabla = false;
                    DbDataReader reader = Command.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            Object objValor = null;
                            blnErrorDato = false;
                            blnErrorCampo = false;
                            blnErrorDatoTemp = false;
                            blnErrorCampoTemp = false;

                            #region Validamos los campos de DocumentoEquipamiento
                            DocumentoEquipamiento = new DocumentoEquipamientoBE();
                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "CODIGO DE KIT", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoEquipamiento.KitSIAE.CodigoGilat = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "SERIE DE KIT", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                            {
                                DocumentoEquipamiento.KitSIAE.SerieKit = Convert.ToString(objValor);
                                DocumentoEquipamiento.KitSIAE.SerieKit = DocumentoEquipamiento.KitSIAE.SerieKit.Substring(DocumentoEquipamiento.KitSIAE.SerieKit.Length - 4);
                            }


                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, true, "NUMERO DE SERIE", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoEquipamiento.SerieEquipamiento = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "CODIGO GILAT", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoEquipamiento.Equipamiento.IdValor = Convert.ToString(objValor);

                            DocumentoEquipamiento.IdEmpresa = "A";//AIO
                            #endregion

                            #region Si no hay errores de campo o de dato intentamos procesar la fila.
                            if (!blnErrorCampo && !blnErrorDato)
                            {
                                try
                                {
                                    if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                    {
                                        DocumentoEquipamiento.UsuarioCreacion = UsuarioCreacion;
                                        DocumentoEquipamientoBL.InsertarDocumentoEquipamientoProceso(DocumentoEquipamiento, baseDatosDA);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    blnErrorDato = true;
                                    file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                }
                            }
                            #endregion
                            intFila++;
                        }
                    }
                    else
                    {
                        file.WriteLine("La tabla EQUIPAMIENTO AIO no tiene registros.");
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    if (ex.GetType().FullName.Equals("System.Data.OleDb.OleDbException") && ((System.Data.OleDb.OleDbException)ex).ErrorCode.Equals(-2147467259))
                    {
                        file.WriteLine("No existe la tabla EQUIPAMIENTO AIO");
                    }
                    else
                    {
                        file.WriteLine(ex.Message);
                    }
                    blnErrorTabla = true;
                }
                #endregion

                #region Validamos la hoja EQUIPAMIENTO DELTRON
                intFila = 2;
                file.WriteLine("");
                file.WriteLine("");
                file.WriteLine("EQUIPAMIENTO DELTRON");
                file.WriteLine("--------------------");
                Command = new OleDbCommand("SELECT * FROM [EQUIPAMIENTO DELTRON$] WHERE [SERIE DE KIT] <> 'SPARE'", conexionExcel);
                try
                {
                    blnErrorTabla = false;
                    DbDataReader reader = Command.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            Object objValor = null;
                            blnErrorDato = false;
                            blnErrorCampo = false;
                            blnErrorDatoTemp = false;
                            blnErrorCampoTemp = false;

                            #region Validamos los campos de DocumentoEquipamiento
                            DocumentoEquipamiento = new DocumentoEquipamientoBE();
                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "CODIGO DE KIT", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoEquipamiento.KitSIAE.CodigoGilat = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "SERIE DE KIT", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                            {
                                DocumentoEquipamiento.KitSIAE.SerieKit = Convert.ToString(objValor);
                                DocumentoEquipamiento.KitSIAE.SerieKit = DocumentoEquipamiento.KitSIAE.SerieKit.Substring(DocumentoEquipamiento.KitSIAE.SerieKit.Length - 4);
                            }


                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "NUMERO DE SERIE", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoEquipamiento.SerieEquipamiento = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "CODIGO PROVEEDOR", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoEquipamiento.Equipamiento.IdValor = Convert.ToString(objValor);

                            DocumentoEquipamiento.IdEmpresa = "D";//DELTRON
                            #endregion

                            #region Si no hay errores de campo o de dato intentamos procesar la fila.
                            if (!blnErrorCampo && !blnErrorDato)
                            {
                                try
                                {
                                    if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                    {
                                        DocumentoEquipamiento.UsuarioCreacion = UsuarioCreacion;
                                        DocumentoEquipamientoBL.InsertarDocumentoEquipamientoProceso(DocumentoEquipamiento, baseDatosDA);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    blnErrorDato = true;
                                    file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                }
                            }
                            #endregion
                            intFila++;
                        }
                    }
                    else
                    {
                        file.WriteLine("La tabla EQUIPAMIENTO DELTRON no tiene registros.");
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    if (ex.GetType().FullName.Equals("System.Data.OleDb.OleDbException") && ((System.Data.OleDb.OleDbException)ex).ErrorCode.Equals(-2147467259))
                    {
                        file.WriteLine("No existe la tabla EQUIPAMIENTO DELTRON");
                    }
                    else
                    {
                        file.WriteLine(ex.Message);
                    }
                    blnErrorTabla = true;
                }
                #endregion

                #region Validamos la hoja KIT CDTEL APURIMAC
                intFila = 2;
                file.WriteLine("");
                file.WriteLine("");
                file.WriteLine("KIT CDTEL APURIMAC");
                file.WriteLine("------------------");
                //Command = new OleDbCommand("SELECT * FROM [KIT CDTEL APURIMAC$] WHERE [CODIGO GILAT] = 'OS6450-BP-D'", conexionExcel);
                Command = new OleDbCommand("SELECT * FROM [KIT CDTEL APURIMAC$] WHERE [CODIGO GILAT] = 'SWITCH-CMP-0006'", conexionExcel);
                try
                {
                    blnErrorTabla = false;
                    DbDataReader reader = Command.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            Object objValor = null;
                            blnErrorDato = false;
                            blnErrorCampo = false;
                            blnErrorDatoTemp = false;
                            blnErrorCampoTemp = false;

                            #region Validamos los campos de DocumentoEquipamiento
                            DocumentoEquipamiento = new DocumentoEquipamientoBE();
                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "NODO DESTINO", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoEquipamiento.Documento.Tarea.NodoIIBBA.IdNodo = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "CODIGO GILAT", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoEquipamiento.Equipamiento.IdValor = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "CODIGO DE KIT DE REFERECIA", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoEquipamiento.KitSIAE.CodigoGilat = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "S/N", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoEquipamiento.SerieEquipamiento = Convert.ToString(objValor);

                            DocumentoEquipamiento.IdEmpresa = "AP";//APURIMAC

                            #endregion

                            #region Si no hay errores de campo o de dato intentamos procesar la fila.
                            if (!blnErrorCampo && !blnErrorDato)
                            {
                                try
                                {
                                    if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                    {
                                        DocumentoEquipamiento.UsuarioCreacion = UsuarioCreacion;
                                        DocumentoEquipamientoBL.InsertarDocumentoEquipamientoProceso(DocumentoEquipamiento, baseDatosDA);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    blnErrorDato = true;
                                    file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                }
                            }
                            #endregion
                            intFila++;
                        }
                    }
                    else
                    {
                        file.WriteLine("La tabla KIT CDTEL APURIMAC no tiene registros.");
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    if (ex.GetType().FullName.Equals("System.Data.OleDb.OleDbException") && ((System.Data.OleDb.OleDbException)ex).ErrorCode.Equals(-2147467259))
                    {
                        file.WriteLine("No existe la tabla KIT CDTEL APURIMAC");
                    }
                    else
                    {
                        file.WriteLine(ex.Message);
                    }
                    blnErrorTabla = true;
                }
                #endregion

                #region Validamos la hoja CDTEL HCVA - AYA
                intFila = 2;
                file.WriteLine("");
                file.WriteLine("");
                file.WriteLine("CDTEL HCVA - AYA");
                file.WriteLine("----------------");
                //Command = new OleDbCommand("SELECT * FROM [CDTEL HCVA - AYA$] WHERE [CODIGO GILAT] = 'OS6450-BP-D'", conexionExcel);
                Command = new OleDbCommand("SELECT * FROM [CDTEL HCVA - AYA$] WHERE [CODIGO GILAT] = 'SWITCH-CMP-0006'", conexionExcel);
                try
                {
                    blnErrorTabla = false;
                    DbDataReader reader = Command.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            Object objValor = null;
                            blnErrorDato = false;
                            blnErrorCampo = false;
                            blnErrorDatoTemp = false;
                            blnErrorCampoTemp = false;

                            #region Validamos los campos de DocumentoEquipamiento
                            DocumentoEquipamiento = new DocumentoEquipamientoBE();
                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "NODO DESTINO", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoEquipamiento.Documento.Tarea.NodoIIBBA.IdNodo = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "CODIGO GILAT", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoEquipamiento.Equipamiento.IdValor = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "CODIGO DE KIT DE REFERECIA", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoEquipamiento.KitSIAE.CodigoGilat = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "S/N", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoEquipamiento.SerieEquipamiento = Convert.ToString(objValor);

                            DocumentoEquipamiento.IdEmpresa = "HA";//HUANCAVELICA Y AYACUCHO

                            #endregion

                            #region Si no hay errores de campo o de dato intentamos procesar la fila.
                            if (!blnErrorCampo && !blnErrorDato)
                            {
                                try
                                {
                                    if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                    {
                                        DocumentoEquipamiento.UsuarioCreacion = UsuarioCreacion;
                                        DocumentoEquipamientoBL.InsertarDocumentoEquipamientoProceso(DocumentoEquipamiento, baseDatosDA);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    blnErrorDato = true;
                                    file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                }
                            }
                            #endregion
                            intFila++;
                        }
                    }
                    else
                    {
                        file.WriteLine("La tabla CDTEL HCVA - AYA no tiene registros.");
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    if (ex.GetType().FullName.Equals("System.Data.OleDb.OleDbException") && ((System.Data.OleDb.OleDbException)ex).ErrorCode.Equals(-2147467259))
                    {
                        file.WriteLine("No existe la tabla CDTEL HCVA - AYA");
                    }
                    else
                    {
                        file.WriteLine(ex.Message);
                    }
                    blnErrorTabla = true;
                }
                #endregion

                #region Validamos la hoja DISTRIBUCION
                intFila = 2;
                file.WriteLine("");
                file.WriteLine("");
                file.WriteLine("DISTRIBUCION");
                file.WriteLine("------------");
                //Command = new OleDbCommand("SELECT * FROM [EQUIPAMIENTO DELTRON$] WHERE [SERIE DE KIT] <> 'SPARE' AND [NUMERO DE SERIE] = '83121611105192'", conexionExcel);
                Command = new OleDbCommand("SELECT * FROM [DISTRIBUCION$]", conexionExcel);
                try
                {
                    blnErrorTabla = false;
                    DbDataReader reader = Command.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            Object objValor = null;
                            blnErrorDato = false;
                            blnErrorCampo = false;
                            blnErrorDatoTemp = false;
                            blnErrorCampoTemp = false;

                            #region Validamos los campos de DocumentoIP
                            DocumentoIP = new DocumentoIPBE();
                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "CODIGO NODO", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoIP.Documento.Tarea.NodoIIBBA.IdNodo = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "IP SYSTEM", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoIP.IPSystem = Convert.ToString(objValor);


                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "RANGO GESTION SEGURIDAD Y ENERGIA (/27)", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoIP.RangoGestionSeguridadEnergia = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "Gateway", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoIP.Gateway = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "MASCARA", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoIP.Mascara = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "IP Reservada", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoIP.IPReservada = Convert.ToString(objValor);

                            DocumentoIP.Documento.Tarea.TipoNodoA.ValorCadena1 = "DISTRIBUCION";

                            #region Controladora inteligente
                            DocumentoIPEquipamiento = new DocumentoIPEquipamientoBE();
                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "Controladora inteligente", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                            {
                                DocumentoIPEquipamiento.IPEquipamiento = Convert.ToString(objValor);
                                DocumentoIPEquipamiento.Equipamiento.IdValor = "CONT-ACC-002";
                                DocumentoIPEquipamiento.DocumentoIP = DocumentoIP;
                                DocumentoIPEquipamiento.UsuarioCreacion = UsuarioCreacion;
                                DocumentoIP.Equipamientos.Add(DocumentoIPEquipamiento);
                            }
                            #endregion

                            #region Cámara 01 PTZ Indoor
                            DocumentoIPEquipamiento = new DocumentoIPEquipamientoBE();
                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "Cámara 01 PTZ Indoor", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                            {
                                DocumentoIPEquipamiento.IPEquipamiento = Convert.ToString(objValor);
                                DocumentoIPEquipamiento.Equipamiento.IdValor = "CMR-IND-0001";
                                DocumentoIPEquipamiento.DocumentoIP = DocumentoIP;
                                DocumentoIPEquipamiento.UsuarioCreacion = UsuarioCreacion;
                                DocumentoIP.Equipamientos.Add(DocumentoIPEquipamiento);
                            }
                            #endregion

                            #region Cámara 02 outdoor
                            DocumentoIPEquipamiento = new DocumentoIPEquipamientoBE();
                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "Cámara 02 outdoor", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                            {
                                DocumentoIPEquipamiento.IPEquipamiento = Convert.ToString(objValor);
                                DocumentoIPEquipamiento.Equipamiento.IdValor = "CMR-OUT-0001";
                                DocumentoIPEquipamiento.DocumentoIP = DocumentoIP;
                                DocumentoIPEquipamiento.UsuarioCreacion = UsuarioCreacion;
                                DocumentoIP.Equipamientos.Add(DocumentoIPEquipamiento);
                            }
                            #endregion

                            #region Grabador de Video NVR
                            DocumentoIPEquipamiento = new DocumentoIPEquipamientoBE();
                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "Grabador de Video NVR", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                            {
                                DocumentoIPEquipamiento.IPEquipamiento = Convert.ToString(objValor);
                                DocumentoIPEquipamiento.Equipamiento.IdValor = "SIS-GRB-0001";
                                DocumentoIPEquipamiento.DocumentoIP = DocumentoIP;
                                DocumentoIPEquipamiento.UsuarioCreacion = UsuarioCreacion;
                                DocumentoIP.Equipamientos.Add(DocumentoIPEquipamiento);
                            }
                            #endregion

                            #region Lector Biométrico
                            DocumentoIPEquipamiento = new DocumentoIPEquipamientoBE();
                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "Lector Biométrico", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                            {
                                DocumentoIPEquipamiento.IPEquipamiento = Convert.ToString(objValor);
                                DocumentoIPEquipamiento.Equipamiento.IdValor = "LEC-ACC-OD-001";
                                DocumentoIPEquipamiento.DocumentoIP = DocumentoIP;
                                DocumentoIPEquipamiento.UsuarioCreacion = UsuarioCreacion;
                                DocumentoIP.Equipamientos.Add(DocumentoIPEquipamiento);
                            }

                            #endregion

                            #endregion

                            #region Si no hay errores de campo o de dato intentamos procesar la fila.
                            if (!blnErrorCampo && !blnErrorDato)
                            {
                                try
                                {
                                    if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                    {
                                        DocumentoIP.UsuarioCreacion = UsuarioCreacion;
                                        DocumentoIP.Documento.Documento.IdValor = "000014";//ACTA DE SEGURIDAD Y DISTRIBUCION
                                        DocumentoIPBL.InsertarDocumentoIPProceso(DocumentoIP, baseDatosDA);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    blnErrorDato = true;
                                    file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                }
                            }
                            #endregion
                            intFila++;
                        }
                    }
                    else
                    {
                        file.WriteLine("La tabla DISTRIBUCION no tiene registros.");
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    if (ex.GetType().FullName.Equals("System.Data.OleDb.OleDbException") && ((System.Data.OleDb.OleDbException)ex).ErrorCode.Equals(-2147467259))
                    {
                        file.WriteLine("No existe la tabla DISTRIBUCION");
                    }
                    else
                    {
                        file.WriteLine(ex.Message);
                    }
                    blnErrorTabla = true;
                }
                #endregion

                #region Validamos la hoja DISTRITAL
                intFila = 2;
                file.WriteLine("");
                file.WriteLine("");
                file.WriteLine("DISTRITAL");
                file.WriteLine("---------");
                //Command = new OleDbCommand("SELECT * FROM [EQUIPAMIENTO DELTRON$] WHERE [SERIE DE KIT] <> 'SPARE' AND [NUMERO DE SERIE] = '83121611105192'", conexionExcel);
                Command = new OleDbCommand("SELECT * FROM [DISTRITAL$]", conexionExcel);
                try
                {
                    blnErrorTabla = false;
                    DbDataReader reader = Command.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            Object objValor = null;
                            blnErrorDato = false;
                            blnErrorCampo = false;
                            blnErrorDatoTemp = false;
                            blnErrorCampoTemp = false;

                            #region Validamos los campos de DocumentoIP
                            DocumentoIP = new DocumentoIPBE();
                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "CODIGO NODO", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoIP.Documento.Tarea.NodoIIBBA.IdNodo = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, true, "IP SYSTEM", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoIP.IPSystem = Convert.ToString(objValor);


                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, true, "RANGO GESTION SEGURIDAD Y ENERGIA (/27)", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoIP.RangoGestionSeguridadEnergia = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, true, "Gateway", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoIP.Gateway = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, true, "MASCARA", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoIP.Mascara = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, true, "IP Reservada", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoIP.IPReservada = Convert.ToString(objValor);

                            DocumentoIP.Documento.Tarea.TipoNodoA.ValorCadena1 = "DISTRITAL";

                            #region Controladora inteligente
                            DocumentoIPEquipamiento = new DocumentoIPEquipamientoBE();
                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, true, "Controladora inteligente", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                            {
                                DocumentoIPEquipamiento.IPEquipamiento = Convert.ToString(objValor);
                                DocumentoIPEquipamiento.Equipamiento.IdValor = "CONT-ACC-002";
                                DocumentoIPEquipamiento.DocumentoIP = DocumentoIP;
                                DocumentoIPEquipamiento.UsuarioCreacion = UsuarioCreacion;
                                DocumentoIP.Equipamientos.Add(DocumentoIPEquipamiento);
                            }
                            #endregion

                            #region Cámara 01 PTZ Indoor
                            DocumentoIPEquipamiento = new DocumentoIPEquipamientoBE();
                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, true, "Cámara 01 PTZ Indoor", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                            {
                                DocumentoIPEquipamiento.IPEquipamiento = Convert.ToString(objValor);
                                DocumentoIPEquipamiento.Equipamiento.IdValor = "CMR-IND-0001";
                                DocumentoIPEquipamiento.DocumentoIP = DocumentoIP;
                                DocumentoIPEquipamiento.UsuarioCreacion = UsuarioCreacion;
                                DocumentoIP.Equipamientos.Add(DocumentoIPEquipamiento);
                            }
                            #endregion

                            #region Cámara 02 outdoor
                            DocumentoIPEquipamiento = new DocumentoIPEquipamientoBE();
                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, true, "Cámara 02 outdoor", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                            {
                                DocumentoIPEquipamiento.IPEquipamiento = Convert.ToString(objValor);
                                DocumentoIPEquipamiento.Equipamiento.IdValor = "CMR-OUT-0001";
                                DocumentoIPEquipamiento.DocumentoIP = DocumentoIP;
                                DocumentoIPEquipamiento.UsuarioCreacion = UsuarioCreacion;
                                DocumentoIP.Equipamientos.Add(DocumentoIPEquipamiento);
                            }
                            #endregion

                            #region Grabador de Video NVR
                            DocumentoIPEquipamiento = new DocumentoIPEquipamientoBE();
                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, true, "Grabador de Video NVR", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                            {
                                DocumentoIPEquipamiento.IPEquipamiento = Convert.ToString(objValor);
                                DocumentoIPEquipamiento.Equipamiento.IdValor = "SIS-GRB-0001";
                                DocumentoIPEquipamiento.DocumentoIP = DocumentoIP;
                                DocumentoIPEquipamiento.UsuarioCreacion = UsuarioCreacion;
                                DocumentoIP.Equipamientos.Add(DocumentoIPEquipamiento);
                            }
                            #endregion

                            #endregion

                            #region Si no hay errores de campo o de dato intentamos procesar la fila.
                            if (!blnErrorCampo && !blnErrorDato)
                            {
                                try
                                {
                                    if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                    {
                                        DocumentoIP.UsuarioCreacion = UsuarioCreacion;
                                        DocumentoIP.Documento.Documento.IdValor = "000013";//ACTA DE SEGURIDAD Y ACCESO
                                        DocumentoIPBL.InsertarDocumentoIPProceso(DocumentoIP, baseDatosDA);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    blnErrorDato = true;
                                    file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                }
                            }
                            #endregion
                            intFila++;
                        }
                    }
                    else
                    {
                        file.WriteLine("La tabla DISTRITAL no tiene registros.");
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    if (ex.GetType().FullName.Equals("System.Data.OleDb.OleDbException") && ((System.Data.OleDb.OleDbException)ex).ErrorCode.Equals(-2147467259))
                    {
                        file.WriteLine("No existe la tabla DISTRITAL");
                    }
                    else
                    {
                        file.WriteLine(ex.Message);
                    }
                    blnErrorTabla = true;
                }
                #endregion

                #region Validamos la hoja INTERMEDIO - TERMINAL
                intFila = 2;
                file.WriteLine("");
                file.WriteLine("");
                file.WriteLine("INTERMEDIO - TERMINAL");
                file.WriteLine("---------------------");
                //Command = new OleDbCommand("SELECT * FROM [EQUIPAMIENTO DELTRON$] WHERE [SERIE DE KIT] <> 'SPARE' AND [NUMERO DE SERIE] = '83121611105192'", conexionExcel);
                Command = new OleDbCommand("SELECT * FROM [INTERMEDIO - TERMINAL$]", conexionExcel);
                try
                {
                    blnErrorTabla = false;
                    DbDataReader reader = Command.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            Object objValor = null;
                            blnErrorDato = false;
                            blnErrorCampo = false;
                            blnErrorDatoTemp = false;
                            blnErrorCampoTemp = false;

                            #region Validamos los campos de DocumentoIP
                            DocumentoIP = new DocumentoIPBE();
                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "CODIGO LOCALIDAD", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoIP.Documento.Tarea.NodoIIBBA.IdNodo = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "Tipo de Nodo", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoIP.Documento.Tarea.TipoNodoA.ValorCadena1 = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "GATEWAY (/23)", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoIP.Gateway = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "MASCARA", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoIP.Mascara = Convert.ToString(objValor);

                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "IP Reservada", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                DocumentoIP.IPReservada = Convert.ToString(objValor);



                            #region Controladora inteligente
                            DocumentoIPEquipamiento = new DocumentoIPEquipamientoBE();
                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "Controladora inteligente", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                            {
                                DocumentoIPEquipamiento.IPEquipamiento = Convert.ToString(objValor);
                                DocumentoIPEquipamiento.Equipamiento.IdValor = "CONT-ACC-002";
                                DocumentoIPEquipamiento.DocumentoIP = DocumentoIP;
                                DocumentoIPEquipamiento.UsuarioCreacion = UsuarioCreacion;
                                DocumentoIP.Equipamientos.Add(DocumentoIPEquipamiento);
                            }
                            #endregion

                            #region Cámara 01 Outdoor
                            DocumentoIPEquipamiento = new DocumentoIPEquipamientoBE();
                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, true, "Cámara 01 Outdoor", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                            {
                                DocumentoIPEquipamiento.IPEquipamiento = Convert.ToString(objValor);
                                DocumentoIPEquipamiento.Equipamiento.IdValor = "CMR-OUT-0001";
                                DocumentoIPEquipamiento.DocumentoIP = DocumentoIP;
                                DocumentoIPEquipamiento.UsuarioCreacion = UsuarioCreacion;
                                DocumentoIP.Equipamientos.Add(DocumentoIPEquipamiento);
                            }
                            #endregion

                            #region Cámara 02 Indoor
                            DocumentoIPEquipamiento = new DocumentoIPEquipamientoBE();
                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, true, "Cámara 02 Indoor", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                            {
                                DocumentoIPEquipamiento.IPEquipamiento = Convert.ToString(objValor);
                                DocumentoIPEquipamiento.Equipamiento.IdValor = "CMR-IND-0002";
                                DocumentoIPEquipamiento.DocumentoIP = DocumentoIP;
                                DocumentoIPEquipamiento.UsuarioCreacion = UsuarioCreacion;
                                DocumentoIP.Equipamientos.Add(DocumentoIPEquipamiento);
                            }
                            #endregion

                            #region Lector RFID 1
                            DocumentoIPEquipamiento = new DocumentoIPEquipamientoBE();
                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, true, "Lector RFID 1", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                            {
                                DocumentoIPEquipamiento.IPEquipamiento = Convert.ToString(objValor);
                                DocumentoIPEquipamiento.Equipamiento.IdValor = "LEC-RFID-0001";
                                DocumentoIPEquipamiento.DocumentoIP = DocumentoIP;
                                DocumentoIPEquipamiento.UsuarioCreacion = UsuarioCreacion;
                                DocumentoIP.Equipamientos.Add(DocumentoIPEquipamiento);
                            }
                            #endregion

                            #region Lector RFID 2
                            DocumentoIPEquipamiento = new DocumentoIPEquipamientoBE();
                            objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, true, "Lector RFID 2", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                            if (blnErrorDatoTemp)
                                blnErrorDato = blnErrorDatoTemp;
                            if (blnErrorCampoTemp)
                                blnErrorCampo = blnErrorCampoTemp;
                            if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                            {
                                DocumentoIPEquipamiento.IPEquipamiento = Convert.ToString(objValor);
                                DocumentoIPEquipamiento.Equipamiento.IdValor = "LEC-RFID-0001";
                                DocumentoIPEquipamiento.DocumentoIP = DocumentoIP;
                                DocumentoIPEquipamiento.UsuarioCreacion = UsuarioCreacion;
                                DocumentoIP.Equipamientos.Add(DocumentoIPEquipamiento);
                            }
                            #endregion

                            #endregion

                            #region Si no hay errores de campo o de dato intentamos procesar la fila.
                            if (!blnErrorCampo && !blnErrorDato)
                            {
                                try
                                {
                                    if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                    {
                                        DocumentoIP.UsuarioCreacion = UsuarioCreacion;
                                        DocumentoIP.Documento.Documento.IdValor = "000013";//ACTA DE SEGURIDAD Y ACCESO
                                        DocumentoIPBL.InsertarDocumentoIPProceso(DocumentoIP, baseDatosDA);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    blnErrorDato = true;
                                    file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                }
                            }
                            #endregion
                            intFila++;
                        }
                    }
                    else
                    {
                        file.WriteLine("La tabla DISTRITAL no tiene registros.");
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    if (ex.GetType().FullName.Equals("System.Data.OleDb.OleDbException") && ((System.Data.OleDb.OleDbException)ex).ErrorCode.Equals(-2147467259))
                    {
                        file.WriteLine("No existe la tabla DISTRITAL");
                    }
                    else
                    {
                        file.WriteLine(ex.Message);
                    }
                    blnErrorTabla = true;
                }
                #endregion

                #region Validamos la hoja ESTUDIO CAMPO APURIMAC
                intFila = 2;
                file.WriteLine("ESTUDIO CAMPO APURIMAC");
                file.WriteLine("----------------------");
                Command = new OleDbCommand("SELECT * FROM [ESTUDIO CAMPO APURIMAC$]", conexionExcel);
                try
                {
                    blnErrorTabla = false;
                    DbDataReader reader = Command.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            if (intFila >= 5)
                            {
                                Object objValor = null;
                                blnErrorDato = false;
                                blnErrorCampo = false;
                                blnErrorDatoTemp = false;
                                blnErrorCampoTemp = false;

                                DocumentoDetalle = new DocumentoDetalleBE();
                                DocumentoDetalle.Documento = new DocumentoBE();
                                DocumentoDetalle.Documento.Documento.IdValor = "000005";//ESTUDIO DE CAMPO

                                #region Nodo
                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "2", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                    DocumentoDetalle.Documento.Tarea.NodoIIBBA.IdNodo = Convert.ToString(objValor);
                                #endregion

                                #region Area Natural Protegida
                                DocumentoDetalle.Campo.IdValor = "000422";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "51", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Nombre Area Natural
                                DocumentoDetalle.Campo.IdValor = "000423";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "52", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Restos Arqueologicos
                                DocumentoDetalle.Campo.IdValor = "000424";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "53", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Tipo de restos arqueologicos
                                DocumentoDetalle.Campo.IdValor = "000425";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "54", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Banco de la Nacion
                                DocumentoDetalle.Campo.IdValor = "000426";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "38", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Agente Banco Nacion
                                DocumentoDetalle.Campo.IdValor = "000427";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "39", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Cantidad
                                DocumentoDetalle.Campo.IdValor = "000428";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "38", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Otros Bancos
                                DocumentoDetalle.Campo.IdValor = "000429";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "39", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Cantidad Otros Bancos
                                DocumentoDetalle.Campo.IdValor = "000430";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "38", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Entidad Importante
                                DocumentoDetalle.Campo.IdValor = "000431";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "37", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Entidad Importante 2
                                DocumentoDetalle.Campo.IdValor = "000506";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "35", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Institucion Educativa
                                DocumentoDetalle.Campo.IdValor = "000432";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "33", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Cantidad Institucion Educativa
                                DocumentoDetalle.Campo.IdValor = "000433";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "32", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Poblacion
                                DocumentoDetalle.Campo.IdValor = "000434";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "16", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region N° DE MUJERES
                                DocumentoDetalle.Campo.IdValor = "000435";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "26", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region N° JÓVENES ENTRE 15 Y 24 AÑOS
                                DocumentoDetalle.Campo.IdValor = "000436";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "24", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region N° DE PERSONAS CON ALGUNA DISCAPACIDAD
                                DocumentoDetalle.Campo.IdValor = "000437";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "28", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region N° DE VIVIENDAS
                                DocumentoDetalle.Campo.IdValor = "000438";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "22", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                            }
                            intFila++;
                        }
                    }
                    else
                    {
                        file.WriteLine("La tabla ESTUDIO CAMPO APURIMAC no tiene registros.");
                    }

                    reader.Close();
                }
                catch (Exception ex)
                {
                    if (ex.GetType().FullName.Equals("System.Data.OleDb.OleDbException") && ((System.Data.OleDb.OleDbException)ex).ErrorCode.Equals(-2147467259))
                    {
                        file.WriteLine("No existe la tabla ESTUDIO CAMPO APURIMAC");
                    }
                    else
                    {
                        file.WriteLine(ex.Message);
                    }
                    blnErrorTabla = true;
                }
                #endregion

                #region Validamos la hoja ESTUDIO CAMPO AYACUCHO
                intFila = 2;
                file.WriteLine("ESTUDIO CAMPO AYACUCHO");
                file.WriteLine("----------------------");
                Command = new OleDbCommand("SELECT * FROM [ESTUDIO CAMPO AYACUCHO$]", conexionExcel);
                try
                {
                    blnErrorTabla = false;
                    DbDataReader reader = Command.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            if (intFila >= 5)
                            {
                                Object objValor = null;
                                blnErrorDato = false;
                                blnErrorCampo = false;
                                blnErrorDatoTemp = false;
                                blnErrorCampoTemp = false;

                                DocumentoDetalle = new DocumentoDetalleBE();
                                DocumentoDetalle.Documento = new DocumentoBE();
                                DocumentoDetalle.Documento.Documento.IdValor = "000005";//ESTUDIO DE CAMPO

                                #region Nodo
                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "2", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                    DocumentoDetalle.Documento.Tarea.NodoIIBBA.IdNodo = Convert.ToString(objValor);
                                #endregion

                                #region Area Natural Protegida
                                DocumentoDetalle.Campo.IdValor = "000422";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "51", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Nombre Area Natural
                                DocumentoDetalle.Campo.IdValor = "000423";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "52", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Restos Arqueologicos
                                DocumentoDetalle.Campo.IdValor = "000424";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "53", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Tipo de restos arqueologicos
                                DocumentoDetalle.Campo.IdValor = "000425";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "54", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Banco de la Nacion
                                DocumentoDetalle.Campo.IdValor = "000426";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "38", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Agente Banco Nacion
                                DocumentoDetalle.Campo.IdValor = "000427";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "39", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Cantidad
                                DocumentoDetalle.Campo.IdValor = "000428";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "38", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Otros Bancos
                                DocumentoDetalle.Campo.IdValor = "000429";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "39", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Cantidad Otros Bancos
                                DocumentoDetalle.Campo.IdValor = "000430";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "38", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Entidad Importante
                                DocumentoDetalle.Campo.IdValor = "000431";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "37", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Entidad Importante 2
                                DocumentoDetalle.Campo.IdValor = "000506";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "35", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Institucion Educativa
                                DocumentoDetalle.Campo.IdValor = "000432";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "33", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Cantidad Institucion Educativa
                                DocumentoDetalle.Campo.IdValor = "000433";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "32", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Poblacion
                                DocumentoDetalle.Campo.IdValor = "000434";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "16", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region N° DE MUJERES
                                DocumentoDetalle.Campo.IdValor = "000435";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "26", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region N° JÓVENES ENTRE 15 Y 24 AÑOS
                                DocumentoDetalle.Campo.IdValor = "000436";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "24", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region N° DE PERSONAS CON ALGUNA DISCAPACIDAD
                                DocumentoDetalle.Campo.IdValor = "000437";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "28", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region N° DE VIVIENDAS
                                DocumentoDetalle.Campo.IdValor = "000438";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "22", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                            }
                            intFila++;
                        }
                    }
                    else
                    {
                        file.WriteLine("La tabla ESTUDIO CAMPO AYACUCHO no tiene registros.");
                    }

                    reader.Close();
                }
                catch (Exception ex)
                {
                    if (ex.GetType().FullName.Equals("System.Data.OleDb.OleDbException") && ((System.Data.OleDb.OleDbException)ex).ErrorCode.Equals(-2147467259))
                    {
                        file.WriteLine("No existe la tabla ESTUDIO CAMPO AYACUCHO");
                    }
                    else
                    {
                        file.WriteLine(ex.Message);
                    }
                    blnErrorTabla = true;
                }
                #endregion

                #region Validamos la hoja ESTUDIO CAMPO HUANCAVELICA
                intFila = 2;
                file.WriteLine("ESTUDIO CAMPO HUANCAVELICA");
                file.WriteLine("--------------------------");
                Command = new OleDbCommand("SELECT * FROM [ESTUDIO CAMPO HUANCAVELICA$]", conexionExcel);
                try
                {
                    blnErrorTabla = false;
                    DbDataReader reader = Command.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            if (intFila >= 5)
                            {
                                Object objValor = null;
                                blnErrorDato = false;
                                blnErrorCampo = false;
                                blnErrorDatoTemp = false;
                                blnErrorCampoTemp = false;

                                DocumentoDetalle = new DocumentoDetalleBE();
                                DocumentoDetalle.Documento = new DocumentoBE();
                                DocumentoDetalle.Documento.Documento.IdValor = "000005";//ESTUDIO DE CAMPO

                                #region Nodo
                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "6", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                    DocumentoDetalle.Documento.Tarea.NodoIIBBA.IdNodo = Convert.ToString(objValor);
                                #endregion

                                #region Area Natural Protegida
                                DocumentoDetalle.Campo.IdValor = "000422";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "50", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Nombre Area Natural
                                DocumentoDetalle.Campo.IdValor = "000423";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "51", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Restos Arqueologicos
                                DocumentoDetalle.Campo.IdValor = "000424";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "52", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Tipo de restos arqueologicos
                                DocumentoDetalle.Campo.IdValor = "000425";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "53", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Banco de la Nacion
                                DocumentoDetalle.Campo.IdValor = "000426";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "37", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Agente Banco Nacion
                                DocumentoDetalle.Campo.IdValor = "000427";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "38", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Cantidad
                                DocumentoDetalle.Campo.IdValor = "000428";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "37", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Otros Bancos
                                DocumentoDetalle.Campo.IdValor = "000429";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "38", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Cantidad Otros Bancos
                                DocumentoDetalle.Campo.IdValor = "000430";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "37", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Entidad Importante
                                DocumentoDetalle.Campo.IdValor = "000431";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "34", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Entidad Importante 2
                                DocumentoDetalle.Campo.IdValor = "000506";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "36", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Institucion Educativa
                                DocumentoDetalle.Campo.IdValor = "000432";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "32", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Cantidad Institucion Educativa
                                DocumentoDetalle.Campo.IdValor = "000433";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "31", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region Poblacion
                                DocumentoDetalle.Campo.IdValor = "000434";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "15", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region N° DE MUJERES
                                DocumentoDetalle.Campo.IdValor = "000435";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "25", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region N° JÓVENES ENTRE 15 Y 24 AÑOS
                                DocumentoDetalle.Campo.IdValor = "000436";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "23", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region N° DE PERSONAS CON ALGUNA DISCAPACIDAD
                                DocumentoDetalle.Campo.IdValor = "000437";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "27", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                                #region N° DE VIVIENDAS
                                DocumentoDetalle.Campo.IdValor = "000438";
                                DocumentoDetalle.Aprobado = true;
                                DocumentoDetalle.Comentario = "Valor insertado de forma automática.";
                                DocumentoDetalle.TipoInsercion = "A";

                                objValor = UtilitarioBL.ValidarDatoReader<String>(reader, intFila, false, "21", out blnErrorCampoTemp, out blnErrorDatoTemp, file);
                                if (blnErrorDatoTemp)
                                    blnErrorDato = blnErrorDatoTemp;
                                if (blnErrorCampoTemp)
                                    blnErrorCampo = blnErrorCampoTemp;
                                if (!blnErrorDatoTemp && !blnErrorCampoTemp)
                                {
                                    DocumentoDetalle.ValorCadena = Convert.ToString(objValor);
                                }

                                #region Si no hay errores de campo o de dato intentamos procesar la fila.
                                if (!blnErrorCampo && !blnErrorDato)
                                {
                                    try
                                    {
                                        if (ddlMetodo.SelectedValue.Equals("000001"))//Insertar
                                        {
                                            DocumentoDetalle.UsuarioCreacion = UsuarioCreacion;
                                            DocumentoDetalleBL.InsertarDocumentoDetalleProceso(DocumentoDetalle, baseDatosDA);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        blnErrorDato = true;
                                        file.WriteLine("Fila " + intFila.ToString() + ": " + ex.Message);
                                    }
                                }
                                #endregion

                                #endregion

                            }
                            intFila++;
                        }
                    }
                    else
                    {
                        file.WriteLine("La tabla ESTUDIO CAMPO HUANCAVELICA no tiene registros.");
                    }

                    reader.Close();
                }
                catch (Exception ex)
                {
                    if (ex.GetType().FullName.Equals("System.Data.OleDb.OleDbException") && ((System.Data.OleDb.OleDbException)ex).ErrorCode.Equals(-2147467259))
                    {
                        file.WriteLine("No existe la tabla ESTUDIO CAMPO HUANCAVELICA");
                    }
                    else
                    {
                        file.WriteLine(ex.Message);
                    }
                    blnErrorTabla = true;
                }
                #endregion

                lnkLog.HRef = rutaVirtualTemporalBE.ValorCadena1 + "/" + Path.GetFileName(((FileStream)(file.BaseStream)).Name);

                #endregion

                baseDatosDA.ConfirmarTransaccion();
            }
            catch (Exception ex)
            {
                baseDatosDA.CancelarTransaccion();
                throw ex;
            }
            finally
            {
                file.Close();
                conexionExcel.Close();

                baseDatosDA.Desconectar();
                baseDatosDA = null;
            }

        }

    }
}
