using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
using System.Globalization;

namespace ExcelFileLoader
{
    class Program
    {
        static void Main(string[] args)
        {
            while (true)
            {
                string filepath = @"C:\Users\admin\Documents\TradeFile2.xlsx";
                DataTable output = exceldata(filepath);
                DataColumnCollection columns = output.Columns;
                Console.Write("Header   : ");
                foreach (DataColumn column in columns)
                {
                    Console.Write(column.ColumnName.PadRight(10,' '));
                    Console.Write("|");
                }
                Console.WriteLine();

                int i = 1;
                foreach (DataRow row in output.Rows)
                {
                    
                    Console.Write("Row " + i.ToString().PadLeft(4,' ') + " : ");
                    i += 1;
                    object[] array = row.ItemArray;
                    foreach (var cell in array)
                    {
                        string cellData = cell == DBNull.Value ? "NULL" : cell.ToString();
                        Console.Write(cellData.PadRight(10, ' '));
                        Console.Write("|");
                    }
                    Console.WriteLine();

                }
                foreach (var v in ExtractColumn(output, "F2"))
                {
                    Console.Write(v == DBNull.Value || v == null ? "NULL" : v.ToString());
                }
                string repeat = Console.ReadLine();
            }
        }

        public static DataTable exceldata(string filePath)
        {
            DataTable dtexcel = new DataTable();
            bool hasHeaders = false;
            string HDR = hasHeaders ? "Yes" : "No";
            string strConn;
            if (filePath.Substring(filePath.LastIndexOf('.')).ToLower() == ".xlsx")
                strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=\"Excel 12.0;IMEX=1;HDR=NO;TypeGuessRows=0;ImportMixedTypes=Text\"";
            else
                strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties=\"Excel 8.0;IMEX=1;HDR=NO;TypeGuessRows=0;ImportMixedTypes=Text\"";
            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            //Looping Total Sheet of Xl File
            /*foreach (DataRow schemaRow in schemaTable.Rows)
            {
            }*/
            //Looping a first Sheet of Xl File
            DataRow schemaRow = schemaTable.Rows[0];
            string sheet = schemaRow["TABLE_NAME"].ToString();
            if (!sheet.EndsWith("_"))
            {
                string query = "SELECT  * FROM ["+ sheet +"];";
                OleDbDataAdapter daexcel = new OleDbDataAdapter(query, conn);
                dtexcel.Locale = CultureInfo.CurrentCulture;
                daexcel.Fill(dtexcel);
            }

            conn.Close();
            return dtexcel;

        }

        public static List<object> ExtractColumn(DataTable input, string columnName)
        {
            return  input.AsEnumerable().Select(r => r.Field<object>(columnName)).ToList();
        }

     
    }
}
