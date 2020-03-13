using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace LibraryReadApp
{
    class Program
    {
        public static DataTable dtResponse = new DataTable();
        public static DataTable dt10K10Q = new DataTable();
        public static DataTable dt8k = new DataTable();
        public static DataTable dtPresentation = new DataTable();
        public static DataTable dtEvents = new DataTable();
        public static DataTable dtNewsPressRelease = new DataTable();
        static void Main(string[] args)
        {
            Console.WriteLine("Retrieving Data from library " + ConfigurationManager.AppSettings.Get("SPOFolder10K10Q") + ".....");
            TenKTenQ tenKTenQ = new TenKTenQ();
            dt10K10Q = tenKTenQ.GetListFileId();
            if (dt10K10Q != null && dt10K10Q.Rows.Count > 0)
                dtResponse.Merge(dt10K10Q);
            Console.WriteLine("Retrieving Data Successful");

            Console.WriteLine("Retrieving Data from library " + ConfigurationManager.AppSettings.Get("SPOFolder8KOther") + ".....");
            EightKOTHERS eightKOTHERS = new EightKOTHERS();
            dt8k = eightKOTHERS.GetListFileId();
            if (dt8k != null && dt8k.Rows.Count > 0)
                dtResponse.Merge(dt8k);
            Console.WriteLine("Retrieving Data Successful");

            Console.WriteLine("Retrieving Data from library " + ConfigurationManager.AppSettings.Get("SPOFolderPresentation") + ".....");
            Presentation presentation = new Presentation();
            dtPresentation = presentation.GetListFileId();
            if (dtPresentation != null && dtPresentation.Rows.Count > 0)
                dtResponse.Merge(dtPresentation);
            Console.WriteLine("Retrieving Data Successful");

            Console.WriteLine("Retrieving Data from library " + ConfigurationManager.AppSettings.Get("SPOFolderEvents") + ".....");
            Events events = new Events();
            dtEvents = events.GetListFileId();
            if (dtEvents != null && dtEvents.Rows.Count > 0)
                dtResponse.Merge(dtEvents);
            Console.WriteLine("Retrieving Data Successful");

            Console.WriteLine("Retrieving Data from library " + ConfigurationManager.AppSettings.Get("SPOFolderNewsPressRelease") + ".....");
            NewsPressRelease newsPressRelease = new NewsPressRelease();
            dtNewsPressRelease = newsPressRelease.GetListFileId();
            if (dtNewsPressRelease != null && dtNewsPressRelease.Rows.Count > 0)
                dtResponse.Merge(dtNewsPressRelease);
            Console.WriteLine("Retrieving Data Successful");

            WriteToCSV(ConfigurationManager.AppSettings.Get("ResponseFilePath") + "Responses.csv", dtResponse);

            Console.ReadLine();
        }

        public static void WriteToCSV(string fileName, DataTable dtDataTable)
        {
            try
            {
                StreamWriter sw = new StreamWriter(fileName, false);

                //for (int i = 0; i < dtDataTable.Rows.Count; i++)
                //{
                //    for (int j = 0; j < dtDataTable.Columns.Count; j++)
                //    {
                //        dtDataTable.Rows[i][j] = dtDataTable.Rows[i][j].ToString().Replace("\n\r", "");
                //        dtDataTable.Rows[i][j] = dtDataTable.Rows[i][j].ToString().Replace("\r\n", "");
                //        //dtDataTable.Rows[i][j] = dtDataTable.Rows[i][j].ToString().Replace(" ", "");
                //    }
                //}

                for (int i = 0; i < dtDataTable.Columns.Count; i++)
                {
                    sw.Write(dtDataTable.Columns[i]);
                    if (i < dtDataTable.Columns.Count - 1)
                    {
                        sw.Write(",");
                    }
                }
                sw.Write(sw.NewLine);
                foreach (DataRow dr in dtDataTable.Rows)
                {
                    for (int i = 0; i < dtDataTable.Columns.Count; i++)
                    {
                        if (!Convert.IsDBNull(dr[i]))
                        {
                            sw.Write(dr[i].ToString());
                        }
                        if (i < dtDataTable.Columns.Count - 1)
                        {
                            sw.Write(",");
                        }
                    }
                    sw.Write(sw.NewLine);
                }
                sw.Close();
                Console.WriteLine("All the responses are written to file path :" + fileName);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error in writing response to CSV :" + ex.Message, "Warning");
            }
        }
    }
}
