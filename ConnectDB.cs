namespace DataAccess.MSACCESS
{
    using System;
    using System.Data;
    using System.Data.OleDb;
    using System.IO;
    using System.Text;
    using System.Windows.Forms;   

    public class common
    {
        private bool b_Connect_Ok = false;
        
        private OleDbConnection sql_DataConnection = new OleDbConnection();

        public common()
        {
            this.sql_DataConnection.ConnectionString = "";
        }

        public void close()
        {
            try
            {
                if (this.sql_DataConnection.State == ConnectionState.Open)
                {
                    this.sql_DataConnection.Close();
                }
            }
            catch (OleDbException exception)
            {
                MessageBox.Show("Lỗi: "+exception.Message);
            }
        }       

        public bool executeNonQuery(string OleDbCommand)
        {
            if (this.open())
            {
                try
                {
                    OleDbCommand command = new OleDbCommand(OleDbCommand);
                    command.Connection = this.sql_DataConnection;
                    command.CommandType = CommandType.Text;
                    command.ExecuteNonQuery();
                    this.close();
                }
                catch (OleDbException exception)
                {
                    MessageBox.Show("Lỗi: "+exception.Message);
                    return false;
                }
                return true;
            }
            return false;
        }

        public bool executeNonQuery(object[] param, object[] value, string storedProcedureName)
        {
            if (this.open())
            {
                try
                {
                    try
                    {
                        OleDbCommand command = new OleDbCommand();
                        command.Connection = this.sql_DataConnection;
                        if (param != null)
                        {
                            for (int i = 0; i < param.Length; i++)
                            {
                                if (value[i] != null)
                                {
                                    command.Parameters.Add(new OleDbParameter("@" + param[i], ToVN6069(value[i])));
                                }
                                else
                                {
                                    command.Parameters.Add(new OleDbParameter("@" + param[i], DBNull.Value));
                                }
                            }
                        }
                        command.CommandType = CommandType.StoredProcedure;
                        command.CommandText = storedProcedureName;
                        command.ExecuteNonQuery();
                    }
                    catch (OleDbException exception)
                    {
                        MessageBox.Show("Lỗi: "+exception.Message);
                        return false;
                    }
                }
                finally
                {
                }
                return true;
            }
            return false;
        }

        public bool executeNonQuery(object[] param, object[] value, OleDbType[] type, string storedProcedureName)
        {
            if (this.open())
            {
                try
                {
                    try
                    {
                        OleDbCommand command = new OleDbCommand();
                        command.Connection = this.sql_DataConnection;
                        if (param != null)
                        {
                            for (int i = 0; i < param.Length; i++)
                            {
                                OleDbParameter parameter;
                                if (value[i] != null)
                                {
                                    parameter = new OleDbParameter("@" + param[i], ToVN6069(value[i]));
                                    parameter.OleDbType = type[i];
                                    command.Parameters.Add(parameter);
                                }
                                else
                                {
                                    parameter = new OleDbParameter("@" + param[i], DBNull.Value);
                                    parameter.OleDbType = type[i];
                                    command.Parameters.Add(parameter);
                                }
                            }
                        }
                        command.CommandType = CommandType.StoredProcedure;
                        command.CommandText = storedProcedureName;
                        command.ExecuteNonQuery();
                    }
                    catch (OleDbException exception)
                    {
                        MessageBox.Show("Lỗi: "+exception.Message);
                        return false;
                    }
                }
                finally
                {
                }
                return true;
            }
            return false;
        }

        public void executeNonQueryWithTran(string OleDbCommand, ref OleDbConnection con, ref OleDbTransaction tran)
        {
            OleDbCommand command = new OleDbCommand(OleDbCommand);
            command.CommandType = CommandType.Text;
            command.Connection = this.sql_DataConnection;
            command.Transaction = tran;
            command.ExecuteNonQuery();
        }

        public void executeNonQueryWithTran(string OleDbCommand, object[] param, object[] value, ref OleDbConnection con, ref OleDbTransaction tran)
        {
            OleDbCommand command = new OleDbCommand(OleDbCommand);
            if (param != null)
            {
                for (int i = 0; i < param.Length; i++)
                {
                    if (value[i] != null)
                    {
                        command.Parameters.AddWithValue("@" + param[i], ToVN6069(value[i]));
                    }
                    else
                    {
                        command.Parameters.AddWithValue("@" + param[i], DBNull.Value);
                    }
                }
            }
            command.CommandType = CommandType.StoredProcedure;
            command.Connection = this.sql_DataConnection;
            command.Transaction = tran;
            command.ExecuteNonQuery();
        }

        public void executeUpdateImage(string OleDbCommand, byte[] arrImage)
        {
            if (this.open())
            {
                try
                {
                    OleDbCommand command = new OleDbCommand(OleDbCommand);
                    command.Connection = this.sql_DataConnection;
                    command.Parameters.Add(new OleDbParameter("@COMPUTER_FILE_CONTENT", OleDbType.VarBinary)).Value = arrImage;
                    command.ExecuteNonQuery();
                    this.close();
                }
                catch (OleDbException exception)
                {
                    MessageBox.Show("Lỗi: "+exception.Message);
                }
            }
        }

        public object executeUpdateScalar(string OleDbCommand)
        {
            if (this.open())
            {
                try
                {
                    OleDbCommand command = new OleDbCommand(OleDbCommand);
                    command.Connection = this.sql_DataConnection;
                    command.CommandType = CommandType.Text;
                    command.ExecuteNonQuery();
                    command = new OleDbCommand("SELECT @@IDENTITY");
                    command.Connection = this.sql_DataConnection;
                    command.CommandType = CommandType.Text;
                    return command.ExecuteScalar();
                }
                catch (OleDbException exception)
                {
                    MessageBox.Show("Lỗi: "+exception.Message);
                    return null;
                }
                finally
                {
                    this.close();
                }
            }
            return null;
        }

        private string GetConfig()
        {
            string str = "";
            str = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + (Application.StartupPath + @"\Database1.mdb");
            return str;
        }

        public OleDbConnection getConnection()
        {
            return this.sql_DataConnection;
        }

        public OleDbDataReader getDataReader(string OleDbCommand)
        {
            if (this.open())
            {
                try
                {
                    OleDbCommand command = new OleDbCommand();
                    command.Connection = this.sql_DataConnection;
                    command.CommandText = OleDbCommand;
                    return command.ExecuteReader();
                }
                catch (OleDbException exception)
                {
                    MessageBox.Show("Lỗi: "+exception.Message);
                }
            }
            return null;
        }

        public DataSet getDataSet(string OleDbCommand)
        {
            if (this.open())
            {
                try
                {
                    DataSet dataSet = new DataSet();
                    OleDbCommandBuilder builder = new OleDbCommandBuilder();
                    OleDbDataAdapter adapter = new OleDbDataAdapter(OleDbCommand, this.sql_DataConnection);
                    adapter.Fill(dataSet);
                    builder.DataAdapter = adapter;
                    this.close();
                    return dataSet;
                }
                catch (OleDbException exception)
                {
                    MessageBox.Show("Lỗi: "+exception.Message);
                }
            }
            return null;
        }

        public DataSet getDataSet(object[] param, object[] value, string storedProcedureName)
        {
            if (this.open())
            {
                try
                {
                    DataSet dataSet = new DataSet();
                    OleDbCommand selectCommand = new OleDbCommand();
                    selectCommand.Connection = this.sql_DataConnection;
                    if (param != null)
                    {
                        for (int i = 0; i < param.Length; i++)
                        {
                            if (value[i] != null)
                            {
                                selectCommand.Parameters.Add(new OleDbParameter("@" + param[i], ToVN6069(value[i])));
                            }
                            else
                            {
                                selectCommand.Parameters.Add(new OleDbParameter("@" + param[i], DBNull.Value));
                            }
                        }
                    }
                    selectCommand.CommandType = CommandType.StoredProcedure;
                    selectCommand.CommandText = storedProcedureName;
                    OleDbCommandBuilder builder = new OleDbCommandBuilder();
                    OleDbDataAdapter adapter = new OleDbDataAdapter(selectCommand);
                    adapter.Fill(dataSet);
                    builder.DataAdapter = adapter;
                    this.close();
                    return dataSet;
                }
                catch (OleDbException exception)
                {
                    MessageBox.Show("Lỗi: "+exception.Message);
                }
            }
            return null;
        }

        public DataTable getDataTable(string OleDbCommand)
        {
            DataSet set = this.getDataSet(OleDbCommand);
            if (set != null)
            {
                return set.Tables[0];
            }
            return null;
        }

        public DataTable getDataTable(object[] param, object[] value, string storedProcedureName)
        {
            DataSet set = this.getDataSet(param, value, storedProcedureName);
            if (set != null)
            {
                return set.Tables[0];
            }
            return null;
        }

        public object getObject(string OleDbCommand)
        {
            if (this.open())
            {
                try
                {
                    OleDbCommand command = new OleDbCommand(OleDbCommand);
                    command.Connection = this.sql_DataConnection;
                    command.CommandType = CommandType.Text;
                    return command.ExecuteScalar();
                }
                catch (OleDbException exception)
                {
                    MessageBox.Show("Lỗi: "+exception.Message);
                    return null;
                }
                finally
                {
                    this.close();
                }
            }
            return null;
        }

        public object getObjectStoredProcedure(object[] param, object[] values, string storedProcedureName)
        {
            if (this.open())
            {
                try
                {
                    OleDbCommand command = new OleDbCommand(storedProcedureName);
                    command.Connection = this.sql_DataConnection;
                    if (param != null)
                    {
                        for (int i = 0; i < param.Length; i++)
                        {
                            if (values[i] != null)
                            {
                                command.Parameters.Add(new OleDbParameter("@" + param[i], values[i]));
                            }
                            else
                            {
                                command.Parameters.Add(new OleDbParameter("@" + param[i], DBNull.Value));
                            }
                        }
                    }
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = storedProcedureName;
                    return command.ExecuteScalar();
                }
                catch (OleDbException exception)
                {
                    MessageBox.Show("Lỗi: "+exception.Message);
                    return null;
                }
                finally
                {
                    this.close();
                }
            }
            return null;
        }

        public DateTime GetServerDateTime()
        {
            string oleDbCommand = "SELECT GetDate() as CurrentDate";
            DataTable table = this.getDataTable(oleDbCommand);
            if ((table != null) && (table.Rows.Count > 0))
            {
                return Convert.ToDateTime(table.Rows[0]["CurrentDate"]);
            }
            return DateTime.Now;
        }

        public bool open()
        {
            this.TestConnect();
            if (this.sql_DataConnection.ConnectionString == "")
            {
                MessageBox.Show("Lỗi kết nối dữ liệu","Thông báo");
                return false;
            }
            try
            {
                if (this.sql_DataConnection.State == ConnectionState.Closed)
                {
                    this.sql_DataConnection.Open();
                }
                return true;
            }
            catch (OleDbException exception)
            {
                MessageBox.Show("Lỗi: "+exception.Message);
            }
            return false;
        }        

        private void TestConnect()
        {
            if (!this.b_Connect_Ok)
            {
                string config = "";
                config = this.GetConfig();
                if (config == "")
                {
                    MessageBox.Show("Lỗi kết nối dữ liệu.", "Thông báo");                    
                }
                else
                {
                    this.sql_DataConnection.ConnectionString = config;
                    try
                    {
                        if (this.sql_DataConnection.State == ConnectionState.Closed)
                        {
                            this.sql_DataConnection.Open();
                        }
                        this.b_Connect_Ok = true;
                    }
                    catch
                    {
                        b_Connect_Ok = false;
                    }
                }
            }
        }

        public bool UpdateDataset(DataTable tbname, string OleDbCommand)
        {            
            tbname = tbname.GetChanges();
            if (tbname == null) return true;
            if (this.open())
            {
                try
                {                    
                    OleDbCommandBuilder builder = new OleDbCommandBuilder();
                    OleDbDataAdapter adapter = new OleDbDataAdapter(OleDbCommand, this.sql_DataConnection);
                    builder.DataAdapter = adapter;
                    adapter.Update(tbname);
                    this.close();
                    return true;
                }
                catch (OleDbException exception)
                {
                    MessageBox.Show("Lỗi: "+exception.Message);
                }
            }
            return false;
        }

        private object ToVN6069(object value)
        {
            if (value.GetType() != typeof(string))
            {
                return value;
            }
            string str = Convert.ToString(value);
            string[] strArray = new string[] { 
                "à", "ả", "ã", "á", "ạ", "ằ", "ẳ", "ẵ", "ắ", "ặ", "\x00e2̀", "\x00e2̉", "\x00e2̃", "\x00e2́", "\x00e2̣", "è", 
                "ẻ", "ẽ", "é", "ẹ", "\x00eà", "\x00eả", "\x00eã", "\x00eá", "\x00eạ", "ò", "ỏ", "õ", "ó", "ọ", "\x00f4̀", "\x00f4̉", 
                "\x00f4̃", "\x00f4́", "\x00f4̣", "ờ", "ở", "ỡ", "ớ", "ợ", "ù", "ủ", "ũ", "ú", "ụ", "ừ", "ử", "ữ", 
                "ứ", "ự", "ì", "ỉ", "ĩ", "í", "ị", "ỳ", "ỷ", "ỹ", "ý", "ỵ", "À", "Ả", "Ã", "Á", 
                "Ạ", "Ằ", "Ẳ", "Ẵ", "Ắ", "Ặ", "\x00c2̀", "\x00c2̉", "\x00c2̃", "\x00c2́", "\x00c2̣", "È", "Ẻ", "Ẽ", "É", "Ẹ", 
                "\x00cà", "\x00cả", "\x00cã", "\x00cá", "\x00cạ", "Ò", "Ỏ", "Õ", "Ó", "Ọ", "\x00d4̀", "\x00d4̉", "\x00d4̃", "\x00d4́", "\x00d4̣", "Ờ", 
                "Ở", "Ỡ", "Ớ", "Ợ", "Ù", "Ủ", "Ũ", "Ú", "Ụ", "Ừ", "Ử", "Ữ", "Ứ", "Ự", "Ì", "Ỉ", 
                "Ĩ", "Í", "Ị", "Ỳ", "Ỷ", "Ỹ", "Ý", "Ỵ"
             };
            string[] strArray2 = new string[] { 
                "\x00e0", "ả", "\x00e3", "\x00e1", "ạ", "ằ", "ẳ", "ẵ", "ắ", "ặ", "ầ", "ẩ", "ẫ", "ấ", "ậ", "\x00e8", 
                "ẻ", "ẽ", "\x00e9", "ẹ", "ề", "ể", "ễ", "ế", "ệ", "\x00f2", "ỏ", "\x00f5", "\x00f3", "ọ", "ồ", "ổ", 
                "ỗ", "ố", "ộ", "ờ", "ở", "ỡ", "ớ", "ợ", "\x00f9", "ủ", "ũ", "\x00fa", "ụ", "ừ", "ử", "ữ", 
                "ứ", "ự", "\x00ec", "ỉ", "ĩ", "\x00ed", "ị", "ỳ", "ỷ", "ỹ", "\x00fd", "ỵ", "\x00c0", "Ả", "\x00c3", "\x00c1", 
                "Ạ", "Ằ", "Ẳ", "Ẵ", "Ắ", "Ặ", "Ầ", "Ẩ", "Ẫ", "Ấ", "Ậ", "\x00c8", "Ẻ", "Ẽ", "\x00c9", "Ẹ", 
                "Ề", "Ể", "Ễ", "Ế", "Ệ", "\x00d2", "Ỏ", "\x00d5", "\x00d3", "Ọ", "Ồ", "Ổ", "Ỗ", "Ố", "Ộ", "Ờ", 
                "Ở", "Ỡ", "Ớ", "Ợ", "\x00d9", "Ủ", "Ũ", "\x00da", "Ụ", "Ừ", "Ử", "Ữ", "Ứ", "Ự", "\x00cc", "Ỉ", 
                "Ĩ", "\x00cd", "Ị", "Ỳ", "Ỷ", "Ỹ", "\x00dd", "Ỵ"
             };
            for (int i = 0; i <= (strArray.Length - 1); i++)
            {
                str = str.Replace(strArray[i], strArray2[i]);
            }
            return str;
        }       
    }
}

