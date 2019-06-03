using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;


namespace sele
{
   public  class Class2
    {
        public string username { get; set; }
        public static string  fetching_data()
        {
            string usrname = "sumakt27@gmail.com";
            string connstring= "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Users\\USER\\Desktop\\23\\data.xlsx;Extended Properties = 'Excel 12.0 Xml;HDR=YES'";
            OleDbConnection conn = new OleDbConnection(connstring);
            String Query = string.Format("select * from[Sheet1$] where username={0}", usrname);
            OleDbCommand cmd = new OleDbCommand(Query,conn);
            conn.Open();
            OleDbDataAdapter da = new OleDbDataAdapter();
            da.SelectCommand = cmd;
            DataTable t = new DataTable();
            da.Fill(t);
            DataSet set = new DataSet();
            da.Fill(set, "mytab");
          //  List<DataRow> list = new List<DataRow>();
          //  da.TableMappings.CopyTo(list, 0);

            string [] ary = new string[20];
            da.TableMappings.CopyTo(ary, 0);

            usrname = ary.ElementAt(0);

            return usrname;
        }
       public void obj()
        {
            Class2 c2 = new Class2();
            //c2.fetching_data();
            Console.ReadLine();
        }
    }
}
