using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.Common;
using System.Configuration;
using System.Reflection;

namespace DataAccess
{
    /// <summary>
    /// Representa la base de datos en el sistema.
    /// Ofrece los métodos de acceso a misma.
    /// </summary>
    public class DBBaseDatos
    {

        #region "Variables Privadas"

        private DbConnection _Conexion = null;
        private DbCommand _Comando = null;
        private DbTransaction _Transaccion = null;
        private DbDataAdapter _Adaptador = null;

        private string _CadenaConexion;
        #endregion

        #region "Variables Privadas Compartidas"


        private static DbProviderFactory _Factory = null;
        #endregion

        #region "Constructores"

        /// <summary>
        /// Crea una instancia del acceso a la base de datos.
        /// </summary>
        public DBBaseDatos()
        {
            Configurar();

        }

        /// <summary>
        /// Configura el acceso a la base de datos para su utilización.
        /// </summary>
        /// <exception cref="BaseDatosException">Si existe un error al cargar la configuración.</exception>
        public void Configurar()
        {
            try
            {
                AppSettingsReader sAsrConfigReader = new AppSettingsReader();
                string proveedor = sAsrConfigReader.GetValue("Proveedor", "".GetType()).ToString();
                this._CadenaConexion = sAsrConfigReader.GetValue("CadenaConexion", "".GetType()).ToString();
                DBBaseDatos._Factory = DbProviderFactories.GetFactory(proveedor);
            }
            catch (ConfigurationException ex)
            {
                //Throw New BaseDatosException("Error al cargar la configuración del acceso a datos.", ex)
                throw new BaseDatosException(ex.Message, ex);
            }
        }


        /// <summary>
        /// Configura el acceso a la base de datos para su utilización.
        /// </summary>
        /// <param name="Proveedor">Cadena que indica el proveedor de servicio (Dirver)</param>
        /// <param name="CadenaConexion">Cadena de conexión</param>
        /// <exception cref="BaseDatosException">Si existe un error al cargar la configuración.</exception>
        public void Configurar(String Proveedor, String CadenaConexion)
        {
            try
            {
                AppSettingsReader sAsrConfigReader = new AppSettingsReader();
                string proveedor = Proveedor;
                this._CadenaConexion = CadenaConexion;
                DBBaseDatos._Factory = DbProviderFactories.GetFactory(proveedor);
            }
            catch (ConfigurationException ex)
            {
                //Throw New BaseDatosException("Error al cargar la configuración del acceso a datos.", ex)
                throw new BaseDatosException(ex.Message, ex);
            }
        }

        #endregion

        #region "Conexion a la Base de Datos"

        /// <summary>
        /// Permite desconectarse de la base de datos.
        /// </summary>
        public void Desconectar()
        {
            if (this._Conexion.State.Equals(ConnectionState.Open))
            {
                this._Conexion.Close();
            }
        }

        /// <summary>
        /// Se concecta con la base de datos.
        /// </summary>
        /// <exception cref="BaseDatosException">Si existe un error al conectarse.</exception> 
        public void Conectar()
        {
            if ((this._Conexion != null))
            {
                if (this._Conexion.State.Equals(ConnectionState.Open))
                {
                    throw new BaseDatosException("La conexión ya se encuentra abierta.");
                }
            }

            try
            {
                if (this._Conexion == null)
                {
                    this._Conexion = _Factory.CreateConnection();
                    this._Conexion.ConnectionString = _CadenaConexion;
                }
                this._Conexion.Open();
            }
            catch (DataException ex)
            {
                throw new BaseDatosException("Error al conectarse. " + ex.Message);
            }
        }

        #endregion

        #region "Creacion de Comandos"

        /// <summary>
        /// Crea un comando en base a una sentencia SQL o Nombre de un Store Procedure.Ejemplo:
        /// <code>SELECT * FROM Tabla WHERE campo1=@campo1, campo2=@campo2 o SP_MI_PROCEDURE</code>
        /// Guarda el comando para el seteo de parámetros y la posterior ejecución.
        /// </summary>
        /// <param name="sentenciaSQL">La sentencia SQL con el formato: SENTENCIA [param = @param,] o SP_MI_PROCEDURE</param>
        /// <param name="TipoComando">El Tipo del comando: TIPO StoredProcedure o Text</param>
        /// <param name="traerParametros">Determina si se van a traer de forma automatica los parametros en caso sea un procedimento</param>
        /// <param name="commandTimeout">Determina la cantidad de tiempo de espera de la ejecución del comando (segundos)</param>
        public void CrearComando(string sentenciaSQL, CommandType TipoComando, Boolean traerParametros = false,Int32 commandTimeout = 900)
        {
            this._Comando = _Factory.CreateCommand();
            this._Comando.Connection = this._Conexion;
            this._Comando.CommandType = TipoComando;
            this._Comando.CommandText = sentenciaSQL;
            this._Comando.CommandTimeout = commandTimeout;

            if (traerParametros)
            {
                if (TipoComando == CommandType.StoredProcedure)
                {
                    if (this._Conexion == null)
                    {
                        throw new BaseDatosException("La conexión no fue inicializada.");
                    }
                    if (!this._Conexion.State.Equals(ConnectionState.Open))
                    {
                        throw new BaseDatosException("La conexión no fue abierta.");
                    }
                    try
                    {
                        DbCommandBuilder cmComando = _Factory.CreateCommandBuilder();
                        MethodInfo mi = cmComando.GetType().GetMethod("DeriveParameters");

                        ///'''''''''''''''''''''''''''''''''''''
                        //Dim miconexion As New Oracle.DataAccess.Client.OracleConnection("Data Source=trondes.mapfreperu.com;Persist Security Info=True;User ID=tron2000;Password=tron2000;")
                        //Dim micomando As New Oracle.DataAccess.Client.OracleCommand("select * from dual", miconexion)
                        //micomando.CommandType = CommandType.Text
                        //miconexion.Open()
                        //Dim adapter As New Oracle.DataAccess.Client.OracleDataAdapter(micomando)

                        //Dim dtTabla As New DataTable
                        //adapter.Fill(dtTabla)

                        //'Oracle.DataAccess.Client.OracleCommandBuilder.DeriveParameters(micomando)

                        //'Dim adapter As New Oracle.DataAccess.Client.OracleDataAdapter("Data Source=trondes.mapfreperu.com;Persist Security Info=True;User ID=tron2000;Password=tron2000;", "gc_k_ws_bcos_mpe.p_bbva")
                        //Dim builder As New Oracle.DataAccess.Client.OracleCommandBuilder(adapter)



                        //'micomando.Parameters = 
                        ///''''''''''''''''''''''''''''''''''''''''



                        //Para asignar el derive parametes sin complicaciones por comenzar una transaccion
                        if ((this._Transaccion != null))
                        {
                            DbConnection cnConexion = null;
                            cnConexion = _Factory.CreateConnection();
                            cnConexion.ConnectionString = _CadenaConexion;
                            cnConexion.Open();
                            this._Comando.Connection = cnConexion;
                            mi.Invoke(null, new object[] { this._Comando });
                            cnConexion.Close();
                            cnConexion.Dispose();
                            this._Comando.Connection = this._Conexion;
                        }
                        else
                        {
                            mi.Invoke(null, new object[] { this._Comando });
                        }
                    }
                    catch (TargetInvocationException ex)
                    {
                        throw new BaseDatosException("Error al crear el comando.", ex);
                    }
                }
            }

            if ((this._Transaccion != null))
            {
                this._Comando.Transaction = this._Transaccion;
            }
        }

        #endregion

        #region "Asignacion de Parametros"

        /// <summary>
        /// Setea un parámetro como nulo del comando creado.
        /// </summary>
        /// <param name="nombre">El nombre del parámetro cuyo valor será nulo.</param>
        public void AsignarParametroNulo(string nombre, Boolean agregar = false, ParameterDirection direccion = ParameterDirection.Input, DbType tipoDB = DbType.String)
        {
            AsignarParametro(nombre, "", "NULL", System.Type.GetType("System.String"), agregar, direccion, tipoDB);
        }

        /// <summary>
        /// Asigna un parámetro de tipo cadena al comando creado.
        /// </summary>
        /// <param name="nombre">El nombre del parámetro.</param>
        /// <param name="valor">El valor del parámetro.</param>
        public void AsignarParametroCadena(string nombre, string valor, Boolean agregar = false, ParameterDirection direccion = ParameterDirection.Input, DbType tipoDB = DbType.String)
        {
            AsignarParametro(nombre, "'", valor, System.Type.GetType("System.String"), agregar, direccion, tipoDB);
        }


        /// <summary>
        /// Asigna un parámetro de tipo cadena al comando creado. Y ademas le asigna una longitud
        /// </summary>
        /// <param name="nombre">El nombre del parámetro.</param>
        /// <param name="valor">El valor del parámetro.</param>
        /// <param name="tamano">El longitud del parámetro.</param>
        public void AsignarParametroCadena(string nombre, string valor, Int32 tamano, Boolean agregar = false, ParameterDirection direccion = ParameterDirection.Input, DbType tipoDB = DbType.String)
        {
            AsignarParametro(nombre, "'", valor, System.Type.GetType("System.String"), agregar, direccion, tipoDB);
            this._Comando.Parameters[nombre].Size = tamano;
        }

        /// <summary>
        /// Asigna un parámetro de tipo entero al comando creado.
        /// </summary>
        /// <param name="nombre">El nombre del parámetro.</param>
        /// <param name="valor">El valor del parámetro.</param>
        public void AsignarParametroEntero(string nombre, int valor, Boolean agregar = false, ParameterDirection direccion = ParameterDirection.Input, DbType tipoDB = DbType.Int32)
        {
            AsignarParametro(nombre, "", valor.ToString(), System.Type.GetType("System.Int32"), agregar, direccion, tipoDB);
        }


        /// <summary>
        /// Asigna un parámetro de tipo entero largo al comando creado.
        /// </summary>
        /// <param name="nombre">El nombre del parámetro.</param>
        /// <param name="valor">El valor del parámetro.</param>
        public void AsignarParametroLong(string nombre, long valor, Boolean agregar = false, ParameterDirection direccion = ParameterDirection.Input, DbType tipoDB = DbType.Int64)
        {
            AsignarParametro(nombre, "", valor.ToString(), System.Type.GetType("System.Int64"), agregar, direccion, tipoDB);
        }


        /// <summary>
        /// Asigna un parámetro de tipo Double al comando creado.
        /// </summary>
        /// <param name="nombre">El nombre del parámetro.</param>
        /// <param name="valor">El valor del parámetro.</param>
        public void AsignarParametroDouble(string nombre, double valor, Boolean agregar = false, ParameterDirection direccion = ParameterDirection.Input, DbType tipoDB = DbType.Double)
        {
            AsignarParametro(nombre, "", valor.ToString(), System.Type.GetType("System.Double"), agregar, direccion, tipoDB);
        }

        /// <summary>
        /// Asigna un parámetro de tipo fecha al comando creado.
        /// </summary>
        /// <param name="nombre">El nombre del parámetro.</param>
        /// <param name="valor">El valor del parámetro.</param>
        public void AsignarParametroFecha(string nombre, DateTime valor, Boolean agregar = false, ParameterDirection direccion = ParameterDirection.Input, DbType tipoDB = DbType.DateTime)
        {
            AsignarParametro(nombre, "'", valor.ToString(), System.Type.GetType("System.DateTime"), agregar, direccion, tipoDB);
        }

        /// <summary>
        /// Asigna un parámetro de tipo array de Bytes al comando creado.
        /// </summary>
        /// <param name="nombre">El nombre del parámetro.</param>
        /// <param name="valor">El valor del parámetro.</param>
        public void AsignarParametroArrayByte(string nombre, byte[] valor, Boolean agregar = false, ParameterDirection direccion = ParameterDirection.Input, DbType tipoDB = DbType.Binary)
        {
            AsignarParametro(nombre, "'", valor, System.Type.GetType("System.Byte[]"), agregar, direccion, tipoDB);
        }

        /// <summary>
        /// Asigna un parámetro de tipo boolean al comando creado.
        /// </summary>
        /// <param name="nombre">El nombre del parámetro.</param>
        /// <param name="valor">El valor del parámetro.</param>
        public void AsignarParametroBoolean(string nombre, Boolean valor, Boolean agregar = false, ParameterDirection direccion = ParameterDirection.Input, DbType tipoDB = DbType.Boolean)
        {
            AsignarParametro(nombre, "", valor.ToString(), System.Type.GetType("System.Boolean"), agregar, direccion, tipoDB);
        }

        /// <summary>
        /// Devuelve un parámetro de tipo Long. Util para parametros de salida de los procedures.
        /// </summary>
        /// <param name="nombre">El nombre del parámetro.</param>
        public long DevolverParametroLong(string nombre)
        {
            //return Convert.ChangeType(this._Comando.Parameters[nombre].Value, System.Type.GetType("System.Int64"));
            return Convert.ToInt64(this._Comando.Parameters[nombre].Value);
            //AsignarParametro(nombre, "'", valor.ToString(), System.Type.GetType("System.DateTime"))
        }

        /// <summary>
        /// Devuelve un parámetro de tipo Double. Util para parametros de salida de los procedures.
        /// </summary>
        /// <param name="nombre">El nombre del parámetro.</param>
        public Double DevolverParametroDouble(string nombre)
        {
            //return Convert.ChangeType(this._Comando.Parameters[nombre].Value, System.Type.GetType("System.Int64"));
            if (this._Comando.Parameters[nombre].Value is System.DBNull)
                return 0;
            else
                return Convert.ToDouble(this._Comando.Parameters[nombre].Value);
            //AsignarParametro(nombre, "'", valor.ToString(), System.Type.GetType("System.DateTime"))
        }

        /// <summary>
        /// Devuelve un parámetro de tipo int. Util para parametros de salida de los procedures.
        /// </summary>
        /// <param name="nombre">El nombre del parámetro.</param>
        public int DevolverParametroEntero(string nombre)
        {
            //return Convert.ChangeType(this._Comando.Parameters[nombre].Value, System.Type.GetType("System.Int32"));
            return Convert.ToInt32(this._Comando.Parameters[nombre].Value);
            //AsignarParametro(nombre, "'", valor.ToString(), System.Type.GetType("System.DateTime"))
        }

        /// <summary>
        /// Devuelve un parámetro de tipo String. Util para parametros de salida de los procedures.
        /// </summary>
        /// <param name="nombre">El nombre del parámetro.</param>
        public String DevolverParametroCadena(string nombre)
        {
            //return Convert.ChangeType(this._Comando.Parameters[nombre].Value, System.Type.GetType("System.Int32"));
            return this._Comando.Parameters[nombre].Value.ToString();
            //AsignarParametro(nombre, "'", valor.ToString(), System.Type.GetType("System.DateTime"))
        }


        /// <summary>
        /// Devuelve un parámetro de tipo Fecha. Util para parametros de salida de los procedures.
        /// </summary>
        /// <param name="nombre">El nombre del parámetro.</param>
        public DateTime DevolverParametroFecha(string nombre)
        {
            //return Convert.ChangeType(this._Comando.Parameters[nombre].Value, System.Type.GetType("System.Int32"));
            return Convert.ToDateTime(this._Comando.Parameters[nombre].Value);
            //AsignarParametro(nombre, "'", valor.ToString(), System.Type.GetType("System.DateTime"))
        }

        /// <summary>
        /// Devuelve un parámetro de tipo Object. Util para parametros de salida de los procedures.
        /// </summary>
        /// <param name="nombre">El nombre del parámetro.</param>
        public Object DevolverParametroObject(string nombre)
        {
            //return Convert.ChangeType(this._Comando.Parameters[nombre].Value, System.Type.GetType("System.Int32"));
            return this._Comando.Parameters[nombre].Value;
            //AsignarParametro(nombre, "'", valor.ToString(), System.Type.GetType("System.DateTime"))
        }

        //public void AsignarParametroFormulario(System.Web.UI.Control control)
        //{
        //    foreach (DbParameter Parametro in this._Comando.Parameters)
        //    {
        //        string strNombreParametro = Parametro.ParameterName.Replace("@", "");
        //        AsignarParametroControl(Parametro.ParameterName.Replace("@", ""), control);
        //    }

        //}


        //    private void AsignarParametroControl(string NombreParametro, System.Web.UI.Control control)
        //{
        //    //Primero buscamos en los controles que no son padres de otros
        //    dynamic rsControlesUnicos = from miControl in control.Controlswhere Strings.Mid(miControl.ID, 5) == NombreParametro && !(miControl.HasControls())miControl;
        //    if (rsControlesUnicos.Count == 1) {
        //        System.Web.UI.WebControls.TextBox miTextBox = rsControlesUnicos(0);
        //        miTextBox.Text = "Control Encontrado";
        //        return;
        //    } else {
        //        //Si no encuentra buscamos en los controles contenedores
        //        dynamic rsControlesContenedores = from miControl in control.Controlswhere miControl.HasControls()miControl;
        //        foreach (System.Web.UI.Control ControlContenedor in rsControlesContenedores) {
        //            AsignarParametroControl(NombreParametro, ControlContenedor);
        //        }
        //    }

        //}

        //Public Sub AsignarParametroFormulario(ByVal formulario As System.Windows.Forms.Control)

        //    'AsignarParametro(nombre, "'", valor.ToString())
        //End Sub

        /// <summary>
        /// Asigna un parámetro al comando creado.
        /// </summary>
        /// <param name="nombre">El nombre del parámetro.</param>
        /// <param name="separador">El separador que será agregado al valor del parámetro.</param>
        /// <param name="valor">El valor del parámetro.</param>
        private void AsignarParametro(String nombre, String separador, Object valor, Type tipo, Boolean agregar = false, ParameterDirection direccion = ParameterDirection.Input, DbType tipoDB = DbType.String)
        {
            if (this._Comando == null)
            {
                throw new BaseDatosException("El comando no ha sido creado.");
            }
            if (this._Comando.CommandType.Equals(CommandType.Text))
            {
                this._Comando.CommandText = this._Comando.CommandText.Replace(nombre, separador + valor.ToString() + separador);
            }
            else if (_Comando.CommandType.Equals(CommandType.StoredProcedure))
            {
                try
                {
                    if (agregar)
                    {
                        DbParameter Parametro = this._Comando.CreateParameter();
                        Parametro.DbType = tipoDB;
                        //Parametro.DbType = DbType.Binary;
                        Parametro.ParameterName = nombre;
                        Parametro.Direction = direccion;
                        if (valor.ToString() == "NULL")
                        {
                            Parametro.Value = Convert.DBNull;
                        }
                        else
                        {
                            Parametro.Value = Convert.ChangeType(valor, tipo);
                        }
                        this._Comando.Parameters.Add(Parametro);
                    }
                    else
                    {
                        if (valor.ToString() == "NULL")
                        {
                            this._Comando.Parameters[nombre].Value = Convert.DBNull;
                        }
                        else
                        {
                            this._Comando.Parameters[nombre].Value = Convert.ChangeType(valor, tipo);
                        }
                    }

                }
                catch (IndexOutOfRangeException ex)
                {
                    throw new BaseDatosException("No existe el nombre del parámetro.", ex);
                }
            }
        }


        #endregion

        #region "Ejecucion del Comando"

        /// <summary>
        /// Ejecuta el comando creado y retorna el resultado de la consulta.
        /// </summary>
        /// <returns>El resultado de la consulta.</returns>
        /// <exception cref="BaseDatosException">Si ocurre un error al ejecutar el comando.</exception>
        public DbDataReader EjecutarConsulta()
        {
            if (this._Conexion == null)
            {
                throw new BaseDatosException("La conexión no fue inicializada.");
            }
            if (!this._Conexion.State.Equals(ConnectionState.Open))
            {
                throw new BaseDatosException("La conexión no fue abierta.");
            }
            try
            {
                return this._Comando.ExecuteReader();
            }
            catch (Exception ex)
            {
                throw new BaseDatosException("Error al ejecutar consulta.", ex);
            }
            //return null;

        }


        /// <summary>
        /// Ejecuta el comando creado y retorna el resultado de la consulta.
        /// </summary>
        /// <returns>El resultado de la consulta.</returns>
        /// <exception cref="BaseDatosException">Si ocurre un error al ejecutar el comando.</exception>
        public DataTable EjecutarConsultaDataTable()
        {
            if (this._Conexion == null)
            {
                throw new BaseDatosException("La conexión no fue inicializada.");
            }
            if (!this._Conexion.State.Equals(ConnectionState.Open))
            {
                throw new BaseDatosException("La conexión no fue abierta.");
            }
            try
            {
                this._Adaptador = _Factory.CreateDataAdapter();
                _Adaptador.SelectCommand = this._Comando;
                DataSet dsDatos = new DataSet();
                this._Adaptador.Fill(dsDatos);
                return dsDatos.Tables[0];
            }
            catch (Exception ex)
            {
                throw new BaseDatosException("Error al ejecutar consulta.", ex);
            }
            //return null;

        }

        /// <summary>
        /// Ejecuta el comando creado y retorna un escalar.
        /// </summary>
        /// <returns>El escalar que es el resultado del comando.</returns>
        /// <exception cref="BaseDatosException">Si ocurre un error al ejecutar el comando.</exception>
        public int EjecutarEscalar()
        {

            if (!this._Comando.CommandType.Equals(CommandType.Text))
            {
                throw new BaseDatosException("El comando debe ser del tipo texto.");
            }
            if (this._Conexion == null)
            {
                throw new BaseDatosException("La conexión no fue inicializada.");
            }
            if (!this._Conexion.State.Equals(ConnectionState.Open))
            {
                throw new BaseDatosException("La conexión no fue abierta.");
            }
            int escalar = 0;
            try
            {
                escalar = int.Parse(this._Comando.ExecuteScalar().ToString());
            }
            catch (InvalidCastException ex)
            {
                throw new BaseDatosException("Error al ejecutar un escalar.", ex);
            }
            catch (Exception ex)
            {
                throw new BaseDatosException("Error al ejecutar un escalar.", ex);
            }
            return escalar;

        }

        /// <summary>
        /// Ejecuta el comando creado.
        /// </summary>
        public void EjecutarComando()
        {
            this._Comando.ExecuteNonQuery();
        }

        #endregion

        #region "Manejo de Transacciones"

        /// <summary>
        /// Comienza una transacción en base a la conexion abierta.
        /// Todo lo que se ejecute luego de esta ionvocación estará 
        /// dentro de una tranasacción.
        /// </summary>
        public void ComenzarTransaccion()
        {
            if (this._Transaccion == null)
            {
                this._Transaccion = this._Conexion.BeginTransaction();
            }
        }

        /// <summary>
        /// Cancela la ejecución de una transacción.
        /// Todo lo ejecutado entre ésta invocación y su 
        /// correspondiente <c>ComenzarTransaccion</c> será perdido.
        /// </summary>
        /// <remarks></remarks>
        public void CancelarTransaccion()
        {
            if ((this._Transaccion != null))
            {
                this._Transaccion.Rollback();
                this._Transaccion = null;
            }
        }

        /// <summary>
        /// Confirma todo los comandos ejecutados entre el <c>ComanzarTransaccion</c>
        /// y ésta invocación.
        /// </summary>
        public void ConfirmarTransaccion()
        {
            if ((this._Transaccion != null))
            {
                this._Transaccion.Commit();
                this._Transaccion = null;
            }
        }

        #endregion

    }
}
