using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Diagnostics;
using RestSharp;
using Newtonsoft.Json;

namespace StockExcelDemo
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {


            
            DateTime mindate = new DateTime(2020, 5, 9, 16, 20, 0);
            DateTime maxdate = new DateTime(2020, 5, 7, 16, 20, 0);
            //double az = (mindate - DateTime.Now).TotalMilliseconds;
            //Console.WriteLine(az / 1000);

            if((maxdate - DateTime.Now).TotalMilliseconds<0|| (mindate - DateTime.Now).TotalMilliseconds > 0)
            {
                //Debug.WriteLine(Properties.Settings.Default.ida);
                //Properties.Settings.Default.ida= "áđâsd";
                Properties.Settings.Default.Save();
                Debug.WriteLine("LOIOASDIOASDIASDASD");
                
            }

            Debug.WriteLine("StockExcel Started");
            //RestClient client = new RestClient("https://iboard.ssi.com.vn");
            //var request = new RestRequest("gateway/graphql", Method.POST, DataFormat.Json);
            //request.AddHeader("Referer", "https://iboard.ssi.com.vn/bang-gia/hose");
            //var obj = JsonConvert.DeserializeObject("{\"operationName\":\"stockRealtimes\",\"variables\":{\"exchange\":\"hose\"},\"query\":\"query stockRealtimes($exchange: String) {\\n  stockRealtimes(exchange: $exchange) {\\n    stockNo\\n    ceiling\\n    floor\\n    refPrice\\n    stockSymbol\\n  }\\n}\\n\"}");
            ////client.RemoteCertificateValidationCallback = (sender, certificate, chain, sslPolicyErrors) => true;
            //request.AddParameter("application/json; charset=utf-8", obj, ParameterType.RequestBody);

            //var result = client.Post(request);
            //Debug.WriteLine(result.Content);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
