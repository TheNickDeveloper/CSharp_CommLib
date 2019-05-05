using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Reflection;
using DataTable = System.Data.DataTable;

namespace SimpleRecon
{
    public class DataHandlingHelper
    {
        //*************************
        // OLEDB connection
        //*************************

        public OleDbConnection SetExcelConnection(string filePath)
        {
            var connStr = $"Provider=Microsoft.ACE.OLEDB.12.0; Data Source={filePath}; Extended Properties=Excel 12.0;";
            return GetConnection(new OleDbConnection(connStr));
        }

        public OleDbConnection SetAs400Connection(string hostName, string userName = "", string password = "")
        {
            var connStr = $"Provider = IBMDA400; Data Source = {hostName}; Persist Security Info = True;";

            if (userName.Length > 0 || password.Length > 0)
                connStr = connStr + $"User ID = {userName}; password = {password};";

            return GetConnection(new OleDbConnection(connStr));
        }

        public OleDbConnection SetAccessConnection(string filePath, string password = "")
        {
            var connStr = $"Provider=Microsoft.ACE.OLEDB.12.0; Data Source={filePath};";

            if (password.Length > 0)
                connStr = connStr + $"Persist Security Info = True; Jet OLEDB:Database Password ={password}; ";
            else
                connStr = connStr + $"Persist Security Info = False;";

            return GetConnection(new OleDbConnection(connStr));
        }

        public OleDbConnection SetSqlServerConnection(string serverName, string dbName, string portname = "1433"
            , string userName = "", string password = "")
        {
            var connStr = $"Provider = SQLOLEDB; Data Source = {serverName},{portname};Intial Catalog = {dbName};";

            if (userName.Length > 0)
                connStr = connStr + $"User ID = {userName}; Password = {password};";

            return GetConnection(new OleDbConnection(connStr));
        }

        private OleDbConnection GetConnection(OleDbConnection connTarget)
        {
            connTarget.Open();
            return connTarget.State == ConnectionState.Open
                ? connTarget : null;
        }

        //*************************
        // Query conversion
        //*************************

        public DataTable ConvertQueryToDataTable(string sqlQuery, OleDbConnection conn, string dtName = "DataTable")
        {
            var adapter = new OleDbDataAdapter(sqlQuery, conn);
            var dataSet = new DataSet();
            adapter.Fill(dataSet, dtName);

            return dataSet.Tables[dtName];
        }

        public List<DataRow> ConvertQueryToStringList(string sqlQuery, OleDbConnection conn, string dtName = "DataTable")
        {
            var dtTarget = ConvertQueryToDataTable(sqlQuery, conn);

            return (from a in dtTarget.AsEnumerable() select a).ToList();
        }

        //*************************
        // Query conversion
        //*************************

        public DataTable ImportCsvAsDataTable(string filePath, bool isIncludeHeader = true)
        {
            var arrLines = File.ReadAllLines(filePath);
            var headerLabels = arrLines[0].Split(',');
            var dtTemp = new DataTable();
            var startRow = 0;

            //header
            if (isIncludeHeader)
            {
                foreach (var headerWord in headerLabels)
                    dtTemp.Columns.Add(new DataColumn(headerWord));

                startRow = 1;
            }
            else
            {
                for (int i = 0; i < headerLabels.Length; i++)
                {
                    dtTemp.Columns.Add(new DataColumn("f" + (i + 1).ToString()));
                    headerLabels[i] = "f" + (i + 1).ToString();
                }
            }

            //body contents
            for (int row = startRow; row < arrLines.Length; row++)
            {
                var dataWords = arrLines[row].Split(',');
                var dataRow = dtTemp.NewRow();
                var columnIndex = 0;

                foreach (var col in headerLabels)
                    dataRow[col] = dataWords[columnIndex++];

                dtTemp.Rows.Add(dataRow);
            }

            return dtTemp;
        }

        public DataTable ConverArrayToDataTable(Array arrTarget, bool isIncludeHeader = true)
        {
            var dtResult = new DataTable();
            var startRow = 0;
            var headerString = "";

            //header
            if (isIncludeHeader)
            {
                for (int i = 0; i < arrTarget.GetLength(1); i++)
                {
                    dtResult.Columns.Add(Convert.ToString(arrTarget.GetValue(0, i)), typeof(string));
                    headerString = headerString + Convert.ToString(arrTarget.GetValue(0, i)) + ",";
                }
                startRow = 1;
            }
            else
            {
                for (int i = 0; i < arrTarget.GetLength(1); i++)
                {
                    dtResult.Columns.Add("f" + (i + 1).ToString());
                    headerString = headerString + "f" + (i + 1).ToString();
                }
            }

            //body contents
            headerString = headerString.Substring(0, headerString.Count() - 1);
            var arrHeader = headerString.Split(',');

            for (int row = startRow; row < arrTarget.GetLength(0); row++)
            {
                var dataRow = dtResult.NewRow();
                var columnIndex = 0;

                foreach (var col in arrHeader)
                    dataRow[col] = arrTarget.GetValue(row, columnIndex++);

                dtResult.Rows.Add(dataRow);
            }

            return dtResult;
        }

        public string[] ImportTextFileAsStringArray(string filePath)
        {
            return File.ReadAllLines(filePath);
        }

        public string[,] ConvertDatatableToStringArray(DataTable dtTarget, bool isIncludeHeader = true)
        {
            var arrResult = new string[dtTarget.Rows.Count + 1, dtTarget.Columns.Count];
            var startRow = 0;

            if (isIncludeHeader)
            {
                //title name
                for (var i = 0; i < dtTarget.Columns.Count; i++)
                    arrResult[0, i] = dtTarget.Columns[i].ColumnName;

                startRow = 1;
            }

            //table contents
            for (var i = 0; i < dtTarget.Rows.Count; i++)
                for (var j = 0; j < dtTarget.Columns.Count; j++)
                    arrResult[i + startRow, j] = dtTarget.Rows[i][j].ToString();

            return arrResult;
        }

        public string[,] ConverArrayToStringArray(Array arrTarget)
        {
            var usedRow = arrTarget.GetLength(0);
            var usedCol = arrTarget.GetLength(1);
            var arrResult = new string[usedRow, usedCol];

            for (int i = 0; i < usedCol; i++)
                for (int j = 0; j < usedRow; j++)
                    arrResult[j, i] = Convert.ToString(arrTarget.GetValue(j + 1, i + 1));

            return arrResult;
        }

        public List<T> BindList<T>(DataTable dt)
        {

            var listTarget = new List<T>();
            var arrPropinfo = typeof(T).GetProperties();

            foreach (DataRow row in dt.Rows)
            {
                var ob = Activator.CreateInstance<T>();

                foreach (var property in typeof(T).GetProperties())
                {
                    var attribute = property.GetCustomAttribute<FeildAttribute>();

                    if (attribute != null)
                        if (dt.Columns.Contains(attribute.Name))
                            property.SetValue(ob, row[dt.Columns[attribute.Name]]);

                    listTarget.Add(ob);
                }
                listTarget.Add(ob);
            }
            return listTarget;
        }
    }

    public class FeildAttribute : Attribute
    {
        public string Name { get; set; }

        public FeildAttribute(string name)
        {
            Name = name;
        }
    }
}


