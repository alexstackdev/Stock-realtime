using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WebSocketSharp;
using System.Diagnostics;

namespace StockExcelDemo
{
    class socketStock
    {
        public string socketUrl;
        public string socketMethod;
        public WebSocket socketIntance;

        public socketStock(string sUrl, string sMethod)
        {
            this.socketUrl = sUrl;
            this.socketMethod = sMethod;
            this.socketIntance = new WebSocket(this.socketUrl, this.socketMethod);
            this.socketIntance.SslConfiguration.EnabledSslProtocols = System.Security.Authentication.SslProtocols.Tls12;
            this.socketIntance.OnOpen += WebSocket_OnOpen;
            this.socketIntance.OnClose += WebSocket_OnClose;
            this.socketIntance.OnError += WebSocket_OnError;
            
        }
        private void WebSocket_OnError(object sender, ErrorEventArgs e)
        {
            Debug.WriteLine("Websocket error");
            Debug.WriteLine(e.Message);
        }

        private void WebSocket_OnClose(object sender, CloseEventArgs e)
        {
            Debug.WriteLine("Websocket closed!");
        }

        private void WebSocket_OnOpen(object sender, EventArgs e)
        {
            Debug.WriteLine("Websocket opened!");
            //this.socketssi.Send("{\"id\":\"3\",\"type\":\"start\",\"payload\":{\"variables\":{\"arr\":[\"hose:21\",\"hose:9\",\"hose:234\",\"hose:3\",\"hose:15\",\"hose:4\",\"hose:19\",\"hose:156\",\"hose:17\",\"hose:10\",\"hose:18\",\"hose:5\",\"hose:13\",\"hose:22\",\"hose:12\",\"hose:7\",\"hose:24\",\"hose:20\",\"hose:190\",\"hose:212\",\"hose:218\",\"hose:219\",\"hose:220\",\"hose:215\",\"hose:217\",\"hose:199\",\"hose:206\",\"hose:195\",\"hose:216\",\"hose:214\",\"hose:210\",\"hose:211\",\"hose:209\",\"hose:221\",\"hose:598\",\"hose:593\",\"hose:601\",\"hose:586\",\"hose:592\",\"hose:589\",\"hose:158\",\"hose:222\",\"hose:605\",\"hose:155\",\"hose:168\",\"hose:181\",\"hose:240\",\"hose:239\",\"hose:241\",\"hose:165\",\"hose:159\",\"hose:223\",\"hose:228\",\"hose:604\",\"hose:142\",\"hose:173\",\"hose:187\",\"hose:230\",\"hose:242\",\"hose:595\",\"hose:572\",\"hose:236\",\"hose:575\",\"hose:588\",\"hose:599\",\"hose:591\",\"hose:170\",\"hose:237\",\"hose:238\",\"hose:244\",\"hose:583\",\"hose:143\",\"hose:186\",\"hose:587\",\"hose:172\",\"hose:174\",\"hose:231\",\"hose:243\",\"hose:590\",\"hose:596\",\"hose:179\",\"hose:574\",\"hose:246\",\"hose:46\",\"hose:48\",\"hose:153\",\"hose:167\",\"hose:245\",\"hose:157\",\"hose:225\",\"hose:581\",\"hose:160\",\"hose:226\",\"hose:603\",\"hose:149\",\"hose:171\",\"hose:232\",\"hose:582\",\"hose:606\",\"hose:580\",\"hose:585\",\"hose:607\",\"hose:227\",\"hose:141\",\"hose:184\",\"hose:224\",\"hose:169\",\"hose:185\",\"hose:164\",\"hose:180\",\"hose:229\",\"hose:140\",\"hose:161\",\"hose:166\",\"hose:183\",\"hose:608\",\"hose:642\",\"hose:647\",\"hose:658\",\"hose:657\",\"hose:101\",\"hose:33\",\"hose:640\",\"hose:655\",\"hose:656\",\"hose:622\",\"hose:641\",\"hose:629\",\"hose:654\",\"hose:632\",\"hose:643\",\"hose:649\",\"hose:631\",\"hose:30\",\"hose:634\",\"hose:635\",\"hose:637\",\"hose:633\",\"hose:651\",\"hose:653\",\"hose:652\",\"hose:650\",\"hose:648\",\"hose:630\",\"hose:644\",\"hose:646\",\"hose:638\",\"hose:796\",\"hose:792\",\"hose:793\",\"hose:795\",\"hose:794\",\"hose:797\",\"hose:979\",\"hose:978\",\"hose:976\",\"hose:53\",\"hose:981\",\"hose:980\",\"hose:973\",\"hose:974\",\"hose:26\",\"hose:984\",\"hose:983\",\"hose:982\",\"hose:49\",\"hose:985\",\"hose:986\",\"hose:235\",\"hose:93\",\"hose:1152\",\"hose:1149\",\"hose:108\",\"hose:1154\",\"hose:1145\",\"hose:1147\",\"hose:1146\",\"hose:1151\",\"hose:1148\",\"hose:1153\",\"hose:233\",\"hose:1364\",\"hose:1380\",\"hose:1370\",\"hose:1291\",\"hose:1379\",\"hose:1292\",\"hose:1337\",\"hose:1339\",\"hose:1382\",\"hose:1366\",\"hose:1386\",\"hose:1348\",\"hose:1369\",\"hose:1378\",\"hose:1383\",\"hose:1384\",\"hose:1336\",\"hose:1381\",\"hose:1376\",\"hose:1354\",\"hose:41\",\"hose:1372\",\"hose:1338\",\"hose:1363\",\"hose:28\",\"hose:1351\",\"hose:1373\",\"hose:1374\",\"hose:58\",\"hose:1385\",\"hose:1318\",\"hose:1377\",\"hose:1375\",\"hose:68\",\"hose:1368\",\"hose:60\",\"hose:69\",\"hose:1371\",\"hose:1422\",\"hose:163\",\"hose:1420\",\"hose:1419\",\"hose:70\",\"hose:1416\",\"hose:1415\",\"hose:1418\",\"hose:1421\",\"hose:1604\",\"hose:1731\",\"hose:1727\",\"hose:1734\",\"hose:1728\",\"hose:1729\",\"hose:100\",\"hose:1738\",\"hose:1733\",\"hose:1891\",\"hose:1888\",\"hose:1889\",\"hose:1893\",\"hose:1898\",\"hose:1899\",\"hose:1900\",\"hose:1890\",\"hose:1894\",\"hose:1896\",\"hose:1895\",\"hose:1897\",\"hose:52\",\"hose:1892\",\"hose:2027\",\"hose:2023\",\"hose:2019\",\"hose:2026\",\"hose:2017\",\"hose:59\",\"hose:2024\",\"hose:2028\",\"hose:2199\",\"hose:2186\",\"hose:2188\",\"hose:2197\",\"hose:145\",\"hose:2195\",\"hose:2196\",\"hose:2193\",\"hose:2185\",\"hose:2198\",\"hose:2187\",\"hose:2200\",\"hose:2191\",\"hose:2340\",\"hose:2339\",\"hose:2549\",\"hose:2570\",\"hose:2575\",\"hose:2573\",\"hose:2567\",\"hose:2552\",\"hose:2546\",\"hose:2560\",\"hose:2572\",\"hose:50\",\"hose:2558\",\"hose:2554\",\"hose:2550\",\"hose:2577\",\"hose:2576\",\"hose:2578\",\"hose:2579\",\"hose:2545\",\"hose:2557\",\"hose:2562\",\"hose:66\",\"hose:2551\",\"hose:2571\",\"hose:2556\",\"hose:2568\",\"hose:2548\",\"hose:2553\",\"hose:2566\",\"hose:2564\",\"hose:2563\",\"hose:2698\",\"hose:2697\",\"hose:2737\",\"hose:2739\",\"hose:2735\",\"hose:2738\",\"hose:2740\",\"hose:2944\",\"hose:2947\",\"hose:2902\",\"hose:2904\",\"hose:2932\",\"hose:2923\",\"hose:2948\",\"hose:2919\",\"hose:2914\",\"hose:2945\"]},\"extensions\":{},\"operationName\":\"stockRealtimeByList\",\"query\":\"subscription stockRealtimeByList($arr: [String]!) {\\n  stockRealtimeByList(arr: $arr) {\\n    data\\n    __typename\\n  }\\n}\\n\"}}");
            //this.socketIntance.Send("{\"id\":\"4\",\"type\":\"start\",\"payload\":{\"variables\":{\"arr\":[\"hose:21\",\"hose:9\"]},\"extensions\":{},\"operationName\":\"stockRealtimeByList\",\"query\":\"subscription stockRealtimeByList($arr: [String]!) {\\n  stockRealtimeByList(arr: $arr) {\\n    data\\n    __typename\\n  }\\n}\\n\"}}");
        }

        public WebSocket getWebSocket()
        {
            return this.socketIntance;

        }


    }
}
