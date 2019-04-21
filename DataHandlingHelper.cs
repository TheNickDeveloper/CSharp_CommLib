using System;
using System.Collections;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using DataTable = System.Data.DataTable;

namespace VstoHelperTest.Helper
{
    public static class DataHandlingHelper
    {
        public static OleDbConnection As400Connection { get; set; }
        public static OleDbConnection ExcelConnection { get; set; }
        public static OleDbConnection AccessConnection { get; set; }
        public static OleDbConnection SqlServerConnection { get; set; }

        //---------------------------------------
        // OleDb Connection 
        //---------------------------------------

        public static void SetExcelConnection(string filePath, string sqlConnectionQuery)
        {
            ExcelConnection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";" 
                + "Extended Properties=\"Excel 12.0;HDR=yes;IMEX=1;\"");

            ExcelConnection?.Open();
        }

        public static void SetAs400Connection(string dbName, string userName = "", string password = "")
        {
            var connectionString = $"Provider - IBMDA400; Data Source = {dbName}; Persist Security Info = True;";
            if (userName.Length > 0 || password.Length > 0)
                connectionString = connectionString + $"User ID = {userName}; Password = {password};";

            As400Connection = new OleDbConnection(connectionString);
            As400Connection?.Open();
        }

        public static void SetAccessConnection(string filePath, string password = "")
        {
            var connectionString = $"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = {filePath};";
            if (password.Length > 0)
                connectionString = connectionString + $"Persist Security Info=True; Jet OLEDB:Database Password = {password};";
            else
                connectionString = connectionString + "Persist Security Info = False;";

            AccessConnection = new OleDbConnection(connectionString);
        }

        public static void SetSqlServerConnection(string serverName, string dbName, string userName = "", string password = "")
        {
            var connectionString = $"Provider = SQLOLEDB; Data Source = {serverName}; Initial Catalog = {dbName};";
            if (userName.Length > 0)
                connectionString = connectionString + $"User ID ={userName}; Password = {password};";

            SqlServerConnection = new OleDbConnection(connectionString);
        }

        //---------------------------------------
        // OleDb Adapter 
        //---------------------------------------

        public static OleDbDataAdapter SetAdapter(string sqlConnectionQuery,OleDbConnection connection)
        {
            return connection.State == ConnectionState.Open
                ? new OleDbDataAdapter(sqlConnectionQuery, connection)
                : null;
        }
        
        public static DataTable ConvertAdapterADataTable(OleDbDataAdapter adapter)
        {
            var dataSet = new DataSet();
            adapter.Fill(dataSet, "DataTable");
            return dataSet.Tables["DataTable"];
        }

        public static ArrayList ConvertAdapterAsArrayList(OleDbDataAdapter adapter)
        {
            var dataSet = new DataSet();
            adapter.Fill(dataSet, "ArrayList");
            var targetDataRow = dataSet.Tables[0].AsEnumerable().ToArray();

            return new ArrayList(targetDataRow);
        }

        //---------------------------------------
        // Other data file import 
        //---------------------------------------

        public static DataTable ImportCsvFilesAsDatatable(string pathFile, bool isIncludeHeader = true)
        {
            var arrLines = File.ReadAllLines(pathFile);
            var headerLables = arrLines[0].Split(',');
            var datatableContainer = new DataTable();
            var startRow = 0;

            if (isIncludeHeader)
            {
                foreach (var headerContents in headerLables)
                    datatableContainer.Columns.Add(new DataColumn(headerContents));

                startRow = 1;
            }
            else
            {
                for (var i = 0; i < headerLables.Length; i++)
                {
                    datatableContainer.Columns.Add(new DataColumn("f" + (i + 1).ToString()));
                    headerLables[i] = "f" + (i + 1).ToString();
                }
            }

            for(var row = startRow; row < arrLines.Length; row++)
            {
                var dataWords = arrLines[row].Split(',');
                var dataRow = datatableContainer.NewRow();
                var columnIndex = 0;

                foreach (var headerWord in headerLables)
                    dataRow[headerWord] = dataWords[columnIndex++];

                datatableContainer.Rows.Add(dataRow);
            }
            return datatableContainer;
        }

        public static DataTable ConvertArrayToDataTable(Array arrTarget, bool isIncludeHeader = true)
        {
            var resultDataTable = new DataTable();
            var startRow = 0;
            var headerString = "";

            if (isIncludeHeader)
            {
                for (var i = 0; i < arrTarget.GetLength(1); i++)
                {
                    resultDataTable.Columns.Add(Convert.ToString(arrTarget.GetValue(0, 1)), typeof(string));
                    headerString = headerString + Convert.ToString(arrTarget.GetValue(0, 1)) + ",";
                }
                startRow = 1;
            }
            else
            {
                for (var i = 0; i < arrTarget.GetLength(1); i++)
                {
                    resultDataTable.Columns.Add($"f{(i + 1)}");
                    headerString = headerString + "f" + (i + 1) + ",";
                }
            }

            headerString = headerString.Substring(0, headerString.Length - 1);
            var headerLables = headerString.Split(',');

            for (var i = startRow; i < arrTarget.GetLength(0); i++)
            {
                var dataRow = resultDataTable.NewRow();
                var columnIndex = 0;

                foreach (var headerWord in headerLables)
                    dataRow[headerWord] = arrTarget.GetValue(i, columnIndex++);

                resultDataTable.Rows.Add(dataRow);
            }
            return resultDataTable;
        }

        public static string[] ImportTxtDataAsStringArray(string pathFile)
        {
            return File.ReadAllLines(pathFile);
        }

        //---------------------------------------
        //Data types converting
        //---------------------------------------

        public static string[,] ConvertDataTableToArray(DataTable datatableTarget, bool isIncludeHeader = true)
        {
            var result = new string[datatableTarget.Rows.Count + 1, datatableTarget.Columns.Count];
            var startRow = 0;

            if (isIncludeHeader)
            {
                for (var i = 0; i < datatableTarget.Columns.Count; i++)
                    result[0, i] = datatableTarget.Columns[i].ColumnName;

                startRow = 1;
            }

            for (var i = 0; i < datatableTarget.Rows.Count; i++)
                    for (var j = 0; j < datatableTarget.Columns.Count; j++)
                        result[i + startRow, j] = datatableTarget.Rows[i][j].ToString();
            
            return result;
        }

        public static string[,] ConvertSystemArrayToStringArray(Array arrTarget)
        {
            var usedRow = arrTarget.GetLength(0);
            var usedCol = arrTarget.GetLength(1);
            var result = new string[usedRow, usedCol];

            for (var i = 0; i < usedCol; i++)
                for (var j = 0; j < usedRow; j++)
                    result[j, i] = Convert.ToString(arrTarget.GetValue(j + 1, i + 1));

            return result;
        }

        public static List<T> BindList<T>(DataTable dt)
        {
            var listTarget = new List<T>();

            foreach(DataRow row in dt.Rows)
            {
                foreach (var property in typeof(T).GetProperties())
                {
                    var attribute = property.GetCustomAttribute<FieldAttribute>();

                    if (attribute != null)
                    {
                        if(dt.Columns.contains(attribute.Name))
                        {
                            property.SetValue(ob,row[dt.Columns[attribute.Name]]);
                        }
                    }
                    listTarget.Add(ob);
                }
            }
            return listTarget;
        }
    }

    public class FeildAttribute:Attribute
    {
        public string Name { get; set; }

        public FeildAttribute (string name)
	    {
            Name = name;
	    }
    }
}


