using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using RestSharp;
using System.Diagnostics;
using Newtonsoft.Json;
//using Newtonsoft.Json.Linq;
using WebSocketSharp;
using System.Net.Http;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using System.Linq;
//using System.Windows.Forms;

//using System.Net.WebSockets;

namespace StockExcelDemo
{
    public partial class Ribbon1
    {
        System.Windows.Forms.Timer timer;
        System.Windows.Forms.Timer timerColor;
        //int autoLoadInterval = 2000;
        WebSocket webSocket = null;
        public int rowLength;
        bool firstLoad = false;
        bool isLoadLable = false;
        bool isAutoLoadLegacy = false;
        int last = 0;


        bool switchAll=false, switchHOSE = false, switchHNX = false, switchUPCOM = false, switchPhaiSinh = false;



        //List<String> rowID = new List<String>();
        Dictionary<String, String> rowID = new Dictionary<string, string>();
        
        object[,] arr;

        int[] num_id_socket = { 0, 0, 61, 6, 7, 4, 5, 2, 3, 42, 43, 52, 22, 23, 24, 25, 26, 27, 44, 46, 47, 54, 0, 0, 0 };
        /* 1
             * 2 Gia 1 ben mua
             * 3 KL 1 ben mua 
             * 4 gia 2 ben mua
             * 5 KL2 ben mua
             * 6 gia 3 ben mua
             * 7 KL3 ben mua 
             * 22 Gia 1 ben ban
             * 23 KL1 ben ban
             * 24 Gia 2 ben ban
             * 25 KL2 ben ban
             * 26 gia 3 ben ban
             * 27 KL 3 ben ban
             * 
             * 
             * 42 Gia
             * 43 KL
             * 44 Cao
             * 45 hose
             * 46 Thap
             * 47 trung binh
             * 52 +/-
             * 54 tong KL
             * 61 TC
             * 75 khac 61
             * 
             */
        public void create_lableName(Application app, int sheetID=1)
        {
            isLoadLable = true;
            Workbook workbook = app.ActiveWorkbook;
            Worksheet worksheet = workbook.Sheets[sheetID];
            //Ck
            Range name_ck = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[2, 1]];
            name_ck.Merge();
            name_ck.BorderAround2();
            name_ck.Value = "CK";

            //Trần
            Range name_tran = worksheet.Range[worksheet.Cells[1, 2], worksheet.Cells[2, 2]];
            name_tran.Merge();
            name_tran.BorderAround2();
            name_tran.Value = "Trần";

            //Sàn
            Range name_san = worksheet.Range[worksheet.Cells[1, 3], worksheet.Cells[2, 3]];
            name_san.Merge();
            name_san.BorderAround2();
            name_san.Value = "Sàn";


            //TC
            Range name_tc = worksheet.Range[worksheet.Cells[1, 4], worksheet.Cells[2, 4]];
            name_tc.Merge();
            name_tc.BorderAround2();
            name_tc.Value = "TC";

            //Bên mua
            Range name_benMua = worksheet.Range[worksheet.Cells[1, 5], worksheet.Cells[1, 10]];
            name_benMua.Merge();
            name_benMua.BorderAround2();
            name_benMua.Value = "Bên mua";

            //Khớp lệnh
            Range name_khopLenh = worksheet.Range[worksheet.Cells[1, 11], worksheet.Cells[1, 13]];
            name_khopLenh.Merge();
            name_khopLenh.BorderAround2();
            name_khopLenh.Value = "Khớp lệnh";

            //Bên bán
            Range name_benBan = worksheet.Range[worksheet.Cells[1, 14], worksheet.Cells[1, 19]];
            name_benBan.Merge();
            name_benBan.BorderAround2();
            name_benBan.Value = "Bên bán";

            //Giá 3
            Range name_benMua_gia3 = worksheet.Cells[2, 5];
            name_benMua_gia3.BorderAround2();
            name_benMua_gia3.Value = "Giá 3";

            //KL 3
            Range name_benMua_KL3 = worksheet.Cells[2, 6];
            name_benMua_KL3.BorderAround2();
            name_benMua_KL3.Value = "KL 3";

            //Giá 2
            Range name_benMua_gia2 = worksheet.Cells[2, 7];
            name_benMua_gia2.BorderAround2();
            name_benMua_gia2.Value = "Giá 2";

            //KL 2
            Range name_benMua_KL2 = worksheet.Cells[2, 8];
            name_benMua_KL2.BorderAround2();
            name_benMua_KL2.Value = "KL 2";

            //Giá 1
            Range name_benMua_gia1 = worksheet.Cells[2, 9];
            name_benMua_gia1.BorderAround2();
            name_benMua_gia1.Value = "Giá 1";

            //KL 1 
            Range name_benMua_KL1 = worksheet.Cells[2, 10];
            name_benMua_KL1.BorderAround2();
            name_benMua_KL1.Value = "KL 1";

            //Giá
            Range name_gia = worksheet.Cells[2, 11];
            name_gia.BorderAround2();
            name_gia.Value = "Giá";

            //KL
            Range name_KL = worksheet.Cells[2, 12];
            name_KL.BorderAround2();
            name_KL.Value = "KL";

            //+/-
            Range name_change = worksheet.Cells[2, 13];
            name_change.BorderAround2();
            name_change.Value = "+/-";

            //Giá 1 - bên bán
            Range name_benBan_gia1 = worksheet.Cells[2, 14];
            name_benBan_gia1.BorderAround2();
            name_benBan_gia1.Value = "Giá 1";

            //KL 1
            Range name_benBan_KL1 = worksheet.Cells[2, 15];
            name_benBan_KL1.BorderAround2();
            name_benBan_KL1.Value = "KL 1";

            //Giá 2
            Range name_benBan_gia2 = worksheet.Cells[2, 16];
            name_benBan_gia2.BorderAround2();
            name_benBan_gia2.Value = "Giá 2";

            //KL 2
            Range name_benBan_KL2 = worksheet.Cells[2, 17];
            name_benBan_KL2.BorderAround2();
            name_benBan_KL2.Value = "KL 2";

            //Giá 3
            Range name_benBan_gia3 = worksheet.Cells[2, 18];
            name_benBan_gia3.BorderAround2();
            name_benBan_gia3.Value = "Giá 3";

            //KL 3
            Range name_benBan_KL3 = worksheet.Cells[2, 19];
            name_benBan_KL3.BorderAround2();
            name_benBan_KL3.Value = "KL 3";

            //Cao
            Range name_cao = worksheet.Range[worksheet.Cells[1, 20], worksheet.Cells[2, 20]];
            name_cao.Merge();
            name_cao.BorderAround2();
            name_cao.Value = "Cao";

            //Thấp
            Range name_thap = worksheet.Range[worksheet.Cells[1, 21], worksheet.Cells[2, 21]];
            name_thap.Merge();
            name_thap.BorderAround2();
            name_thap.Value = "Thấp";

            //TB
            Range name_TB = worksheet.Range[worksheet.Cells[1, 22], worksheet.Cells[2, 22]];
            name_TB.Merge();
            name_TB.BorderAround2();
            name_TB.Value = "TB";

            //Tổng KL
            Range name_tongKL = worksheet.Range[worksheet.Cells[1, 23], worksheet.Cells[2, 23]];
            name_tongKL.Merge();
            name_tongKL.BorderAround2();
            name_tongKL.Value = "Tổng KL";

            //ĐTNN
            Range name_DTNN = worksheet.Range[worksheet.Cells[1, 24], worksheet.Cells[1, 26]];
            name_DTNN.Merge();
            name_DTNN.BorderAround2();
            name_DTNN.Value = "ĐTNN";

            //NN mua
            Range name_nnMua = worksheet.Cells[2, 24];
            name_nnMua.BorderAround2();
            name_nnMua.Value = "NN mua";


            //NN bán
            Range name_nnBan = worksheet.Cells[2, 25];
            name_nnBan.BorderAround2();
            name_nnBan.Value = "NN bán";

            //Dư
            Range name_du = worksheet.Cells[2, 26];
            name_du.BorderAround2();
            name_du.Value = "Dư";

            //toàn bộ tên cột
            Range name_all = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[2, 26]];
            name_all.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            name_all.VerticalAlignment = XlHAlign.xlHAlignCenter;
        }


        private void checkAndCreateSheets(Application app)
        {
            Workbook wb = app.ActiveWorkbook;
            if (!wb.Sheets[1].Name == "HOSE")
            {
                
            }

        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            Debug.WriteLine(DateTime.Now);
            
        }



        public void initValue(Application app, string exchange="hose", int sheetID=1)
        {
            var watch = System.Diagnostics.Stopwatch.StartNew();
            Workbook workbook = app.ActiveWorkbook;

            if (!isLoadLable)
            {
                create_lableName((Application)Marshal.GetActiveObject("Excel.Application"), 1);
                isLoadLable = true;
            }


            if (sheetID == 4) return;



            if (workbook.Sheets[2] == null) Debug.WriteLine("áđá");


            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12 | System.Net.SecurityProtocolType.Tls11 | System.Net.SecurityProtocolType.Tls;
            RestClient client = new RestClient("https://iboard.ssi.com.vn");
            var request = new RestRequest("gateway/graphql", Method.POST, DataFormat.Json);
            request.AddHeader("Referer", "https://iboard.ssi.com.vn/bang-gia/");
            string qr = "{\"operationName\":\"stockRealtimes\",\"variables\":{\"exchange\":\"" + exchange + "\"},\"query\":\"query stockRealtimes($exchange: String) {\\n  stockRealtimes(exchange: $exchange) {\\n    stockNo\\n    ceiling\\n    floor\\n    refPrice\\n    stockSymbol\\n    stockType\\n    exchange\\n    matchedPrice\\n    matchedVolume\\n    priceChange\\n    priceChangePercent\\n    highest\\n    avgPrice\\n    lowest\\n    nmTotalTradedQty\\n    best1Bid\\n    best1BidVol\\n    best2Bid\\n    best2BidVol\\n    best3Bid\\n    best3BidVol\\n    best1Offer\\n    best1OfferVol\\n    best2Offer\\n    best2OfferVol\\n    best3Offer\\n    best3OfferVol\\n    buyForeignQtty\\n    buyForeignValue\\n    sellForeignQtty\\n    sellForeignValue\\n    currentBidQty\\n    currentOfferQty\\n    remainForeignQtty\\n}\\n}\\n\"}";


            var obj = JsonConvert.DeserializeObject(qr);
            
            request.AddParameter("application/json; charset=utf-8", obj, ParameterType.RequestBody);
            
            Worksheet worksheet = workbook.Sheets[2];

            client.PostAsync(request, (res, hanl) =>
            {

                try
                {
                    var result = res;
                    Debug.WriteLine(result.Content);
                    JObject jObject = JObject.Parse(result.Content);
                    rowLength = jObject["data"]["stockRealtimes"].Count();
                    arr = new object[jObject["data"]["stockRealtimes"].Count(), 26];
                    //String[] colName = { "stockSymbol", "ceiling", "floor", "refPrice", "best3Bid", "best3BidVol" , "best2Bid" , "best2BidVol", "best1Bid", "best1BidVol", "matchedPrice",      "matchedVolume", "priceChange", "best1Offer",
                    //    "best1OfferVol","best2Offer",
                    //    "best2OfferVol","best3Offer",
                    //    "best3OfferVol","highest","lowest","avgPrice","nmTotalTradedQty","buyForeignQtty","sellForeignQtty","remainForeignQtty"};
                    for (var i = 3; i < jObject["data"]["stockRealtimes"].Count() + 3; i++)
                    {
                        if (!firstLoad) rowID.Add(jObject["data"]["stockRealtimes"][i - 3]["stockSymbol"].ToString(), i.ToString());

                        JObject oneRow = (JObject)jObject["data"]["stockRealtimes"][i - 3];
                        //string z =  jObject["data"]["stockRealtimes"][i - 3]["stockSymbol"].ToString()+"|"+i.ToString();
                        //Debug.WriteLine(jObject["data"]["stockRealtimes"][i - 3].ToString());

                        arr[i - 3, 0] = jObject["data"]["stockRealtimes"][i - 3]["stockSymbol"].ToString();               //0
                        arr[i - 3, 1] = Double.Parse(oneRow["ceiling"].ToString()) / 1000;
                        arr[i - 3, 2] = Double.Parse(oneRow["floor"].ToString()) / 1000;
                        arr[i - 3, 3] = Double.Parse(oneRow["refPrice"].ToString()) / 1000;
                        arr[i - 3, 4] = Double.Parse(oneRow["best3Bid"].ToString()) / 1000;
                        arr[i - 3, 5] = Double.Parse(oneRow["best3BidVol"].ToString()) / 10;  //5
                        arr[i - 3, 6] = Double.Parse(oneRow["best2Bid"].ToString()) / 1000;
                        arr[i - 3, 7] = Double.Parse(oneRow["best2BidVol"].ToString()) / 10;
                        arr[i - 3, 8] = Double.Parse(oneRow["best1Bid"].ToString()) / 1000;
                        arr[i - 3, 9] = Double.Parse(oneRow["best1BidVol"].ToString()) / 10;
                        arr[i - 3, 10] = Double.Parse(oneRow["matchedPrice"].ToString()) / 1000;    //10
                        arr[i - 3, 11] = Double.Parse(oneRow["matchedVolume"].ToString()) ;
                        if (oneRow["priceChange"].ToString() == "") arr[i - 3, 12] = 0;
                        else arr[i - 3, 12] = Double.Parse(oneRow["priceChange"].ToString()) / 1000;
                        arr[i - 3, 13] = Double.Parse(oneRow["best1Offer"].ToString()) / 1000;
                        arr[i - 3, 14] = Double.Parse(oneRow["best1OfferVol"].ToString()) / 10;
                        arr[i - 3, 15] = Double.Parse(oneRow["best2Offer"].ToString()) / 1000;
                        arr[i - 3, 16] = Double.Parse(oneRow["best2OfferVol"].ToString()) / 10;
                        arr[i - 3, 17] = Double.Parse(oneRow["best3Offer"].ToString()) / 1000;
                        arr[i - 3, 18] = Double.Parse(oneRow["best3OfferVol"].ToString()) / 10;
                        arr[i - 3, 19] = Double.Parse(oneRow["highest"].ToString()) / 1000;
                        arr[i - 3, 20] = Double.Parse(oneRow["lowest"].ToString()) / 1000;
                        arr[i - 3, 21] = Double.Parse(oneRow["avgPrice"].ToString()) / 1000;
                        if (oneRow["nmTotalTradedQty"].ToString() != "") arr[i - 3, 22] = Double.Parse(oneRow["nmTotalTradedQty"].ToString()) / 10;
                        else arr[i - 3, 22] = 0;
                        if (oneRow["buyForeignQtty"].ToString() != "") arr[i - 3, 23] = Double.Parse(oneRow["buyForeignQtty"].ToString()) ;
                        else arr[i - 3, 23] = 0;
                        if (oneRow["sellForeignQtty"].ToString() != "") arr[i - 3, 24] = Double.Parse(oneRow["sellForeignQtty"].ToString()) ;
                        else arr[i - 3, 24] = 0;
                        if (oneRow["remainForeignQtty"].ToString() != "") arr[i - 3, 25] = Double.Parse(oneRow["remainForeignQtty"].ToString()) / 10;
                        else arr[i - 3, 25] = 0;

                    }
                    if (!firstLoad)
                    {
                        firstLoad = true;
                    }
                    Range v = worksheet.Range[worksheet.Cells[3, 1], worksheet.Cells[jObject["data"]["stockRealtimes"].Count()+2, 26]];
                    v.Value2 = arr;
                    v.Borders.LineStyle = Microsoft.Office.Interop.Excel.Constants.xlSolid;

                    Debug.WriteLine("Loaded ", jObject["data"]["stockRealtimes"].Count().ToString(), "row");
                }
                catch (Exception ee)
                {
                    Debug.WriteLine(ee.Message);
                }
                watch.Stop();
                var elapsedMs = watch.ElapsedMilliseconds;
                Debug.WriteLine(Double.Parse(elapsedMs.ToString()) / 1000, " s");
            });

            



        }


        private void btnTaoKhung_Click(object sender, RibbonControlEventArgs e)
        {
            create_lableName((Application)Marshal.GetActiveObject("Excel.Application"));
        }

        private void btnInitValue_Click(object sender, RibbonControlEventArgs e)
        {
            initValue((Application)Marshal.GetActiveObject("Excel.Application"));
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            Debug.WriteLine(DateTime.UtcNow);
            initValue((Application)Marshal.GetActiveObject("Excel.Application"));
        }

        private void btnAutoSwitch_Click(object sender, RibbonControlEventArgs e)
        {
            create_lableName((Application)Marshal.GetActiveObject("Excel.Application"));
            if (webSocket != null)
            {

                btnAutoSwitch.Label = "Bật auto";
                lableAutoStatus.Label = "Cập nhật Tự Động: Tắt";
                webSocket.CloseAsync();
                webSocket = null;
                if (timer != null)
                {
                    timer.Stop();
                    timer = null;

                }
                
            }
            else
            {
                var watch = System.Diagnostics.Stopwatch.StartNew();
                create_lableName((Application)Marshal.GetActiveObject("Excel.Application"));
                initValue((Application)Marshal.GetActiveObject("Excel.Application"));
                btnAutoSwitch.Label = "Tắt auto";
                lableAutoStatus.Label = "Cập nhật Tự Động: Bật";
                webSocket = new socketStock("wss://iboard.ssi.com.vn/realtime/graphql", "graphql-ws").getWebSocket();
                
                webSocket.OnMessage += WebSocket_OnMessage;
                webSocket.OnOpen += WebSocket_OnOpen;
                webSocket.Connect();
                if (timer != null)
                {
                    timer.Stop();
                    timer = null;

                }

                if (timerColor != null)
                {
                    timerColor.Stop();
                    timerColor = null;
                }

                System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12 | System.Net.SecurityProtocolType.Tls11 | System.Net.SecurityProtocolType.Tls;
                RestClient client = new RestClient("https://iboard.ssi.com.vn");
                var request = new RestRequest("gateway/graphql", Method.POST, DataFormat.Json);
                request.AddHeader("Referer", "https://iboard.ssi.com.vn/bang-gia/hose");
                var obj = JsonConvert.DeserializeObject("{\"operationName\":\"stockRealtimes\",\"variables\":{\"exchange\":\"hose\"},\"query\":\"query stockRealtimes($exchange: String) {\\n  stockRealtimes(exchange: $exchange) {\\n    stockNo\\n}\\n}\\n\"}");

                request.AddParameter("application/json; charset=utf-8", obj, ParameterType.RequestBody);
                Workbook workbook = ((Application)Marshal.GetActiveObject("Excel.Application")).ActiveWorkbook;
                Worksheet worksheet = workbook.Sheets[1];

                client.PostAsync(request, (res, hanl) =>
                {
                    try
                    {
                        var result = res;
                        Debug.WriteLine(result.Content.Length);
                        JObject jObject = JObject.Parse(result.Content);
                        rowLength = jObject["data"]["stockRealtimes"].Count();
                        String stockNo = "{\"id\":\"4\",\"type\":\"start\",\"payload\":{\"variables\":{\"arr\":[";
                        for (var i = 0; i < jObject["data"]["stockRealtimes"].Count(); i++)
                        {
                            stockNo += "\"" + jObject["data"]["stockRealtimes"][i]["stockNo"].ToString() + "\"";
                            if (i != jObject["data"]["stockRealtimes"].Count() - 1) stockNo += ",";
                            //Debug.WriteLine(jObject["data"]["stockRealtimes"][i]["stockNo"].ToString());
                        }
                        stockNo+= "]},\"extensions\":{},\"operationName\":\"stockRealtimeByList\",\"query\":\"subscription stockRealtimeByList($arr: [String]!) {\\n  stockRealtimeByList(arr: $arr) {\\n    data\\n    __typename\\n  }\\n}\\n\"}}";
                        //string zz = "{\"id\":\"4\",\"type\":\"start\",\"payload\":{\"variables\":{\"arr\":[\"hose:21\"]},\"extensions\":{},\"operationName\":\"stockRealtimeByList\",\"query\":\"subscription stockRealtimeByList($arr: [String]!) {\\n  stockRealtimeByList(arr: $arr) {\\n    data\\n    __typename\\n  }\\n}\\n\"}}";
                        ////Debug.WriteLine(stockNo);
                        webSocket.SendAsync(stockNo, (z) =>
                        {
                            if (z)
                            {
                                Debug.WriteLine("Send ok");
                            }
                            else Debug.WriteLine("Send error");
                        });
                        //webSocket.SendAsync(zz, (z) =>
                        //{
                        //    if (z)
                        //    {
                        //        Debug.WriteLine("Send ok");
                        //    }
                        //    else Debug.WriteLine("Send error");
                        //});

                    }
                    catch (Exception ee)
                    {
                        Debug.WriteLine(ee.Message);
                    }


                    watch.Stop();
                    var elapsedMs = watch.ElapsedMilliseconds;
                    Debug.WriteLine(Double.Parse(elapsedMs.ToString()) / 1000, " s");
                });

                timerColor = new System.Windows.Forms.Timer();
                timerColor.Tick += new EventHandler((ze,er)=> {
                    Range v = worksheet.Range[worksheet.Cells[3, 1], worksheet.Cells[rowID.Count+2, 26]];
                    v.Interior.Color = System.Drawing.Color.White;
                });
                
                timerColor.Interval = 2000; // in miliseconds
                timerColor.Start();
            }
        }

        private void WebSocket_OnOpen(object sender, EventArgs e)
        {
            Debug.WriteLine("eadsasdas");
        }

        public void WebSocket_OnMessage(object sender, MessageEventArgs e)
        {
            //Debug.WriteLine("asd", (Int32.Parse(DateTime.Now.ToString()) - last).ToString());
            //last = Int32.Parse(DateTime.Now.ToString());

            try
            {
                
                var watch = System.Diagnostics.Stopwatch.StartNew();
                //Debug.WriteLine(e.Data);

                JObject jObject = JObject.Parse(e.Data);

                if (jObject["type"].ToString() == "data" && jObject["id"].ToString() == "4")
                {
                    
                    string res = jObject["payload"]["data"]["stockRealtimeByList"]["data"].ToString();
                    
                    //Debug.WriteLine("");
                    res.Trim();

                    Workbook workbook = ((Application)Marshal.GetActiveObject("Excel.Application")).ActiveWorkbook;
                    Worksheet worksheet = workbook.ActiveSheet;
                    //Range range = worksheet.Range[worksheet.Cells[3, 1], worksheet.Cells[rowLength + 3, 1]];


                    string[] after = res.Split('|');


                    
                    
                    if (after[1] == "") return;
                        int idRow = Int32.Parse(rowID[after[1]]);
                        
                    if (idRow != null)
                    {
                        Debug.WriteLine(last);
                        last++;
                        for (var i = 0; i < num_id_socket.Length; i++)
                        {

                            if (num_id_socket[i] == 0)
                            {
                                //Debug.Write("_    _");
                            }

                            else
                            {
                                int id = num_id_socket[i];
                                if (id == 2 || id == 4 || id == 6 || id == 22 || id == 24 || id == 26 || id == 42 || id == 52 || id == 46 || id == 47 || id == 61 || id == 44)
                                {
                                    //if (Double.Parse(worksheet.Cells[idRow, i + 2].Value.ToString()) > Double.Parse(after[num_id_socket[i]].ToString()) / 1000)
                                    //{
                                    //    worksheet.Cells[idRow, i + 2].Interior.Color = XlRgbColor.rgbSkyBlue;
                                    //    worksheet.Cells[idRow, i + 2].Font.Color = XlRgbColor.rgbSkyBlue;
                                    //}
                                    //else
                                    //{
                                    //    worksheet.Cells[idRow, i + 2].Interior.Color = XlRgbColor.rgbRed;
                                    //    worksheet.Cells[idRow, i + 2].Font.Color = XlRgbColor.rgbRed;
                                    //}

                                    worksheet.Cells[idRow, i + 2].Value = Double.Parse(after[num_id_socket[i]].ToString()) / 1000;
                                   // Debug.WriteLine(id, "ok");

                                }
                                else if (id == 3 || id == 5 || id == 7 || id == 23 || id == 25 || id == 27 || id == 43 || id == 54)
                                {
                                   // Debug.WriteLine(id, "ok");
                                    worksheet.Cells[idRow, i + 2].Value = Double.Parse(after[num_id_socket[i]].ToString()) / 10;
                                    //if (Double.Parse(worksheet.Cells[idRow, i + 2].Value.ToString()) > Double.Parse(after[num_id_socket[i]].ToString()) / 10)
                                    //{
                                    //    worksheet.Cells[idRow, i + 2].Interior.Color = XlRgbColor.rgbSkyBlue;
                                    //    worksheet.Cells[idRow, i + 2].Font.Color = XlRgbColor.rgbSkyBlue;
                                    //}
                                    //else
                                    //{
                                    //    worksheet.Cells[idRow, i + 2].Interior.Color = XlRgbColor.rgbRed;
                                    //    worksheet.Cells[idRow, i + 2].Font.Color = XlRgbColor.rgbRed;
                                    //}

                                }
                                else
                                {
                                   // Debug.WriteLine(id, "ok");
                                    worksheet.Cells[idRow, i + 2] = after[num_id_socket[i]].ToString();
                                    //worksheet.Cells[idRow, i + 2].Interior.Color = XlRgbColor.rgbAqua;
                                }



                            }

                        }

                    }


                    //Debug.WriteLine("");
                }
                watch.Stop();
                var elapsedMs = watch.ElapsedMilliseconds;
                //Debug.WriteLine(Double.Parse(elapsedMs.ToString()) / 1000, " s");
            }
            catch (Exception z)
            {
                Debug.WriteLine(z.Message);
            }
            
        }

        private void btnSendWs_Click(object sender, RibbonControlEventArgs e)
        {
            
            if (webSocket!=null&& edtWs.Text.ToString()!="") webSocket.Send(edtWs.Text.ToString());
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            if(!isAutoLoadLegacy)
            {
                if (!isLoadLable)
                {
                    create_lableName((Application)Marshal.GetActiveObject("Excel.Application"));
                    isLoadLable = true;
                }
                initValue((Application)Marshal.GetActiveObject("Excel.Application"));
                timer = new System.Windows.Forms.Timer();
                timer.Tick += new EventHandler(timer1_Tick);
                btnLoadLegacy.Label = "Tự load mỗi 2s: ON";
                timer.Interval = 2000; // in miliseconds
                timer.Start();
                isAutoLoadLegacy = true;
            }
            else
            {
                isAutoLoadLegacy = false;
                timer.Stop();
                timer = null;
                btnLoadLegacy.Label = "Tự load mỗi 2s: OFF";
                
            }
           
        }

        private void btnHOSE_Click(object sender, RibbonControlEventArgs e)
        {
            switchHOSE = !switchHOSE;
            if (switchHOSE) btnHOSE.Label = "HOSE: OFF";
            else btnHOSE.Label = "HOSE: ON";

        }

        private void btnUPCOM_Click(object sender, RibbonControlEventArgs e)
        {
            switchUPCOM = !switchUPCOM;
            if (switchUPCOM) btnUPCOM.Label = "HOSE: OFF";
            else btnUPCOM.Label = "HOSE: ON";
        }

        private void btnPhaiSinh_Click(object sender, RibbonControlEventArgs e)
        {
            switchPhaiSinh = !switchPhaiSinh;
            if (switchPhaiSinh) btnPhaiSinh.Label = "Phái sinh: OFF";
            else btnPhaiSinh.Label = "Phái sinh: ON";
        }


        private void button1_Click_1(object sender, RibbonControlEventArgs e)
        {
            //Workbook workbook = ((Application)Marshal.GetActiveObject("Excel.Application")).ActiveWorkbook;
            //if (workbook.Sheets["asdasdasdad"]) Debug.WriteLine("exist");

            //else Debug.WriteLine("asdasd");
            System.Windows.Forms.Form f = new Form1();
            f.Show();

        }

        private void dropDown1_SelectionChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void dropDown15_SelectionChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void edtHOSEOpen1_TextChanged(object sender, RibbonControlEventArgs e)
        {
            
        }

        private void btnAutoIncrease_Click(object sender, RibbonControlEventArgs e)
        {
            if (timer != null)
            {
                
                timer.Interval = timer.Interval + 500;
                string a = (Double.Parse(timer.Interval.ToString()) / 1000).ToString();
                btnLoadLegacy.Label = "Tự load mỗi " +a+ " s: ON";
            }
        }

        private void btnAutoDecrease_Click(object sender, RibbonControlEventArgs e)
        {
            if (timer != null&&timer.Interval>1499)
            {
                timer.Interval = timer.Interval - 500;
                string a = (Double.Parse(timer.Interval.ToString()) / 1000).ToString();
                btnLoadLegacy.Label = "Tự load mỗi " +a+ " s: ON";
            }
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnSwitchAll_Click(object sender, RibbonControlEventArgs e)
        {
            switchAll = !switchAll;
            if (switchAll)
            {
                switchHNX = switchHOSE = switchPhaiSinh = switchUPCOM = true;
                btnSwitchAll.Label = "Tất Cả: ON";
            }
            else
            {
                switchHNX = switchHOSE = switchPhaiSinh = switchUPCOM = false;
                btnSwitchAll.Label = "Tất Cả: OFF";
            }
        }
    }
}
