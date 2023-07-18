using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using ExcelDataReader;

namespace Lab5Games.ExcelTool
{
    public static class DataReader
    {
        public static DataRowCollection ReadExcel(Stream fileStream, int sheetIndex = 0)
        {
            IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(fileStream);
            DataSet dataSet = reader.AsDataSet();
            DataTable dataTable = dataSet.Tables[sheetIndex].Rows
                .Cast<DataRow>()
                .Where(row => !row.ItemArray.All(f => f is DBNull ||
                    string.IsNullOrEmpty(f as string ?? f.ToString())))
                .CopyToDataTable();

            return dataTable.Rows;
        }

        public static DataRowCollection ReadCsv(Stream fileStream)
        {
            using(StreamReader reader = new StreamReader(fileStream))
            {
                DataTable dataTable = new DataTable();

                List<string> lines = new List<string>();
                string line = reader.ReadLine();

                while (line != null)
                {
                    lines.Add(line);

                    line = reader.ReadLine();
                }

                // columns
                int len = lines[0].Split(',').Length;
                for (int i = 0; i < len; i++)
                {
                    dataTable.Columns.Add("c_" + i);
                }

                // rows
                for (int i = 0; i < lines.Count; i++)
                {
                    DataRow row = dataTable.NewRow();
                    row.ItemArray = lines[i].Split(',');
                    dataTable.Rows.Add(row);
                }

                dataTable = dataTable.Rows
                    .Cast<DataRow>()
                    .Where(row => !row.ItemArray.All(f => f is DBNull ||
                        string.IsNullOrEmpty(f as string ?? f.ToString())))
                    .CopyToDataTable();

                return dataTable.Rows;
            }
        }

        public static DataRowCollection ReadCsv(string content)
        {
            using(StringReader reader = new StringReader(content))
            {
                DataTable dataTable = new DataTable();

                List<string> lines = new List<string>();
                string line = reader.ReadLine();
                
                while(line != null)
                {
                    lines.Add(line);

                    line = reader.ReadLine();
                }

                // columns
                int len = lines[0].Split(',').Length;
                for(int i=0; i<len; i++)
                {
                    dataTable.Columns.Add("c_" + i);
                }

                // rows
                for(int i=0; i<lines.Count; i++)
                {
                    DataRow row = dataTable.NewRow();
                    row.ItemArray = lines[i].Split(',');
                    dataTable.Rows.Add(row);
                }

                dataTable = dataTable.Rows
                    .Cast<DataRow>()
                    .Where(row => !row.ItemArray.All(f => f is DBNull ||
                        string.IsNullOrEmpty(f as string ?? f.ToString())))
                    .CopyToDataTable();

                return dataTable.Rows;  
            }
        }
    }
}
