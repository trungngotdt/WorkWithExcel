using ExcelDataReader;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace WorkWithExecl
{
    public class ExcelReader
    {
        private String address;

        public String Address
        {
            get { return address; }
            set { address = value; }
        }
        public ExcelReader()
        {

        }
        public ExcelReader(string ar)
        {
            this.Address = ar;
        }
        public List<object> GetData()
        {
            var fs = File.Open(Address,FileMode.Open,FileAccess.Read);
            IExcelDataReader reader = ExcelReaderFactory.CreateReader(fs);
            List<object> list = new List<object>();
            object value;
            while (reader.Read())
            {
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    value = reader.GetValue(i);
                    list.Add(value);
                    list.Add("\t");
                }
                list.Add("\n");
            }
            return list;
        }

    }
}
