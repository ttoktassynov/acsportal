using System;
using System.ComponentModel;
using System.Collections;
using System.Diagnostics;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

public class ACSDB : IDisposable
{
    // connection to data source
    private SqlConnection con;
    private string connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["SQLSRVDEVTESTConnectionToACS"].ConnectionString;
    //private string connectionString2 = System.Configuration.ConfigurationManager.ConnectionStrings["SQLSRVDEVTESTConnectionToKMG"].ConnectionString;
    //private string connectionString3 = System.Configuration.ConfigurationManager.ConnectionStrings["SQLSRVDEVConnectionToKMG"].ConnectionString;
    /// <summary>
    /// Run stored procedure.
    /// </summary>
    /// <param name="procName">Name of stored procedure.</param>
    /// <returns>Stored procedure return value.</returns>

    public ACSDB(string connstr)
    {
        connectionString = connstr;
    }
    public ACSDB()
    {

    }
    
    public int RunProc(string procName)
    {
        SqlCommand cmd = CreateCommand(procName, null);
        cmd.ExecuteNonQuery();
        this.Close();
        return (int)cmd.Parameters["ReturnValue"].Value;
    }

    /// <summary>
    /// Run stored procedure.
    /// </summary>
    /// <param name="procName">Name of stored procedure.</param>
    /// <param name="prams">Stored procedure params.</param>
    /// <returns>Stored procedure return value.</returns>
    public int RunProc(string procName, SqlParameter[] prams)
    {
        SqlCommand cmd = CreateCommand(procName, prams);
        cmd.ExecuteNonQuery();
        this.Close();
        return (int)cmd.Parameters["ReturnValue"].Value;
    }

    /// <summary>
    /// Run stored procedure.
    /// </summary>
    /// <param name="procName">Name of stored procedure.</param>
    /// <param name="dataReader">Return result of procedure.</param>
    public void RunProc(string procName, out SqlDataReader dataReader)
    {
        SqlCommand cmd = CreateCommand(procName, null);
        dataReader = cmd.ExecuteReader(System.Data.CommandBehavior.CloseConnection);
    }

    /// <summary>
    /// Run stored procedure.
    /// </summary>
    /// <param name="procName">Name of stored procedure.</param>
    /// <param name="prams">Stored procedure params.</param>
    /// <param name="dataReader">Return result of procedure.</param>
    public void RunProc(string procName, SqlParameter[] prams, out SqlDataReader dataReader)
    {
        SqlCommand cmd = CreateCommand(procName, prams);
        dataReader = cmd.ExecuteReader(System.Data.CommandBehavior.CloseConnection);
    }

    /// <summary>
    /// Run stored procedure.
    /// </summary>
    /// <param name="procName">Name of stored procedure.</param>
    /// <param name="prams">Stored procedure params.</param>
    /// <param name="dataTable">Return result of procedure.</param>
    public void RunProc(string procName, SqlParameter[] prams, out DataTable dataTable)
    {
        DataTable dt = new DataTable();
        SqlCommand cmd = CreateCommand(procName, prams);
        SqlDataAdapter da = new SqlDataAdapter(cmd);
        int rows = da.Fill(dt);
        dataTable = dt;
    }


    /// <summary>
    /// Create command object used to call stored procedure.
    /// </summary>
    /// <param name="procName">Name of stored procedure.</param>
    /// <param name="prams">Params to stored procedure.</param>
    /// <returns>Command object.</returns>
    private SqlCommand CreateCommand(string procName, SqlParameter[] prams)
    {
        // make sure connection is open
        Open();


        SqlCommand cmd = new SqlCommand(procName, con);
        cmd.CommandType = CommandType.StoredProcedure;
        cmd.CommandTimeout = 3600 * 2 * 24;
        // add proc parameters
        if (prams != null)
        {
            foreach (SqlParameter parameter in prams)
                cmd.Parameters.Add(parameter);
        }

        // return param
        cmd.Parameters.Add(
            new SqlParameter("ReturnValue", SqlDbType.Int, 4,
            ParameterDirection.ReturnValue, false, 0, 0,
            string.Empty, DataRowVersion.Default, null));

        return cmd;
    }

    /// <summary>
    /// Open the connection.
    /// </summary>
    private void Open()
    {
        // open connection
        if (con == null)
        {
            //con = new SqlConnection(ConfigurationSettings.AppSettings["KMG_Report.Properties.Settings.SQLSRVDEVConnection"]);
            //con = new SqlConnection("Data Source=SQLSRVDEV;Initial Catalog=STDB;Integrated Security = True; User Id = tima ; password = Connect_021!");
            //con = new SqlConnection(ConfigurationSettings.AppSettings["SQLSRVDEVConnection"]);
            con = new SqlConnection(connectionString);
            con.Open();
        }
        if (con.State == ConnectionState.Closed)
        {
            con.Open();
        }
    }

    /// <summary>
    /// Close the connection.
    /// </summary>
    public void Close()
    {
        if (con != null)
            con.Close();
    }

    /// <summary>
    /// Release resources.
    /// </summary>
    public void Dispose()
    {
        // make sure connection is closed
        if (con != null)
        {
            con.Dispose();
            con = null;
        }
    }

    /// <summary>
    /// Make input param.
    /// </summary>
    /// <param name="ParamName">Name of param.</param>
    /// <param name="DbType">Param type.</param>
    /// <param name="Size">Param size.</param>
    /// <param name="Value">Param value.</param>
    /// <returns>New parameter.</returns>
    public SqlParameter MakeInParam(string ParamName, SqlDbType DbType, int Size, object Value)
    {
        return MakeParam(ParamName, DbType, Size, ParameterDirection.Input, Value);
    }

    /// <summary>
    /// Make input param.
    /// </summary>
    /// <param name="ParamName">Name of param.</param>
    /// <param name="DbType">Param type.</param>
    /// <param name="Size">Param size.</param>
    /// <returns>New parameter.</returns>
    public SqlParameter MakeOutParam(string ParamName, SqlDbType DbType, int Size)
    {
        return MakeParam(ParamName, DbType, Size, ParameterDirection.Output, null);
    }

    /// <summary>
    /// Make stored procedure param.
    /// </summary>
    /// <param name="ParamName">Name of param.</param>
    /// <param name="DbType">Param type.</param>
    /// <param name="Size">Param size.</param>
    /// <param name="Direction">Parm direction.</param>
    /// <param name="Value">Param value.</param>
    /// <returns>New parameter.</returns>
    public SqlParameter MakeParam(string ParamName, SqlDbType DbType, Int32 Size,
        ParameterDirection Direction, object Value)
    {
        SqlParameter param;

        if (Size > 0)
            param = new SqlParameter(ParamName, DbType, Size);
        else
            param = new SqlParameter(ParamName, DbType);

        param.Direction = Direction;
        if (!(Direction == ParameterDirection.Output && Value == null))
            param.Value = Value;

        return param;
    }

    public DataTable GetDataTable(string Sql)
    {
        SqlDataAdapter da = new SqlDataAdapter(Sql, connectionString);

        DataTable dt = new DataTable();

        int rows = da.Fill(dt);

        return dt;
    }
}

