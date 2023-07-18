using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using ExcelDataReader;

namespace Lab5Games.ExcelTool
{
    public class ExcelReader : IDisposable
    {
        bool _disposed = false;

        Stream _fileStream = null;
        DataSet _dataSet = null;

        public ExcelReader(Stream fileStream)
        {
            _fileStream = fileStream;

            IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(_fileStream);
            _dataSet = reader.AsDataSet();    
        }

        public DataRowCollection Read(int sheetIndex)
        {
            DataTable dt = _dataSet.Tables[sheetIndex].Rows
                .Cast<DataRow>()
                .Where(row => !row.ItemArray.All(f => f is DBNull ||
                    string.IsNullOrEmpty(f as string ?? f.ToString())))
                .CopyToDataTable();

            return dt.Rows;
        }

        public void Dispose()
        {
            Dispose(true);
        }

        protected virtual void Dispose(bool disposing)
        {
            try
            {
                if((!_disposed && disposing))
                {
                    if(_fileStream != null)
                        _fileStream.Close();

                    if (_dataSet != null)
                        _dataSet.Dispose();
                }
            }
            finally
            {
                if(!_disposed)
                {
                    _fileStream = null;
                    _dataSet = null;
                }

                _disposed = true;
            }
        }
    }
}
