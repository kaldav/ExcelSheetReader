using System;
using System.Data;

namespace ExcelSheetReader
{
    class Program
    {

        //Microsoft Access Database Engine 2010 Redistributable needs to be installed
        static void Main(string[] args)
        {
            CreateNewOutputRows();
            Console.ReadLine();
        }

        static void CreateNewOutputRows()
        {
            string fileName = @"C:\Users\david.kalacska\Documents\CALPADS\codesetsv13-20210719.xlsx"; //TODO: should come from parameter
            string sheetName = "CALPADS Code Sets";
            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";

            using (System.Data.OleDb.OleDbConnection xlConn = new System.Data.OleDb.OleDbConnection(connectionString))
            {
                xlConn.Open();
                System.Data.OleDb.OleDbCommand xlCmd = xlConn.CreateCommand();
                xlCmd.CommandText = "Select * from [" + sheetName + "$A3:D50000]"; // works only until 50000 lines
                xlCmd.CommandType = CommandType.Text;
                using (System.Data.OleDb.OleDbDataReader rdr = xlCmd.ExecuteReader())
                {
                    while (rdr.Read())
                    {
                        Console.WriteLine(Int32.Parse(rdr[0].ToString())); //Id
                        Console.WriteLine(rdr.GetString(1)); //CodeSetName
                        Console.WriteLine(rdr.GetString(2)); //CodedValue
                        Console.WriteLine(rdr.GetString(3)); //Name


                        //The first 4 columns are static and added to every row
                        //Output0Buffer.AddRow();
                        //Output0Buffer.UniqueID = Int32.Parse(rdr[0].ToString());
                        //Output0Buffer.Year = Int32.Parse(rdr[1].ToString());
                        //Output0Buffer.ReportingWave = rdr.GetString(2);
                        //Output0Buffer.SubmissionDate = rdr.GetString(3);
                        //Output0Buffer.Question = rdr.GetName(i);
                        //Output0Buffer.Answer = rdr.GetString(i);

                    }
                }
                xlConn.Close();
            }
        }
    }
}
