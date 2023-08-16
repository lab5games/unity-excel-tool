using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace Lab5Games.ExcelTool
{
    public class CsvReader  
    {
        
        public CsvReader()
        {
        }

        public DataRowCollection Read(Stream fileStream)
        {
            using(StreamReader reader = new StreamReader(fileStream))
            {
                DataTable dt = new DataTable();

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
                    dt.Columns.Add("c_" + i);
                }

                // rows
                for(int i=0; i<lines.Count; i++)
                {
                    DataRow row = dt.NewRow();
                    row.ItemArray = lines[i].Split(',');

                    bool hasData = !row.ItemArray.All(f => f is DBNull ||
                        string.IsNullOrEmpty(f as string ?? f.ToString()));

                    if (hasData) dt.Rows.Add(row);
                }

                fileStream.Close();

                return dt.Rows;
            }
        }

        public DataRowCollection Read(string content)
        {
            using(StringReader reader = new StringReader(content))
            {
                DataTable dt = new DataTable();

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
                    dt.Columns.Add("c_" + i);
                }

                // rows
                for (int i = 0; i < lines.Count; i++)
                {
                    DataRow row = dt.NewRow();
                    row.ItemArray = Split(lines[i]);

                    bool hasData = !row.ItemArray.All(f => f is DBNull ||
                        string.IsNullOrEmpty(f as string ?? f.ToString()));

                    if (hasData) dt.Rows.Add(row);
                }

                return dt.Rows;
            }
        }

        private string[] Split(string line, char separator = ',')
        {
            const char ArrayLeft = '[';
            const char ArrayRight = ']';

            bool inArray = false;
            var token = "";
            var result = new List<string>();

            for (int i = 0; i < line.Length; i++)
            { 
                var c = line[i];

                if(inArray)
                {
                    if(c == ArrayRight)
                    {
                        inArray = false;
                        token += ArrayRight;
                    }
                    else
                    {
                        token += c;
                    }
                }
                else
                {
                    if(c == ArrayLeft)
                    {
                        inArray = true;
                        token += ArrayLeft;
                    }
                    else if (c == separator)
                    {
                        result.Add(token);
                        token = "";
                    }
                    else
                    {
                        token += c;
                    }
                }
            }

            return result.ToArray();
        }
    }
}
