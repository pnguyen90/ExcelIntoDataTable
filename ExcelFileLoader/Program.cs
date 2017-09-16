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
                string filepath = @"C:\Users\T-420\Documents\TradeFile.xlsx";
                DataTable output = exceldata(filepath);
                DataColumnCollection columns = output.Columns;
                Console.Write("Row 1 : ");
                foreach (DataColumn column in columns)
                {
                    Console.Write(column.ColumnName);
                    Console.Write("|");
                }
                Console.WriteLine();

                int i = 2;
                foreach (DataRow row in output.Rows)
                {
                    Console.Write("Row " + i + " : ");
                    i += 1;
                    object[] array = row.ItemArray;
                    foreach (var cell in array)
                    {
                        Console.Write(cell.ToString());
                        Console.Write("|");
                    }
                    Console.WriteLine();
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

     
    }
}
