using System;
using System.Activities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;

using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization.Json;

namespace Microsoft.Crm.Sdk.Samples
{

    public sealed partial class CustomActivityForOpp : CodeActivity
    {
        // Azure ML Web API と JSON でやりとりするクラス
        public class Rootobject
        {
            public Results Results { get; set; }
        }

        public class Results
        {
            public Output1 output1 { get; set; }
        }

        public class Output1
        {
            public string type { get; set; }
            public Value value { get; set; }
        }

        public class Value
        {
            public string[] ColumnNames { get; set; }
            public string[] ColumnTypes { get; set; }
            public string[][] Values { get; set; }
        }


        /// <summary>
        /// 営業案件の確度をAzureMLで予測し、結果をCRMに出力する
        /// </summary>
        protected override void Execute(CodeActivityContext executionContext)
        {
            string webapikey = this.InputApikey.Get(executionContext);
            string uri = this.InputUri.Get(executionContext);
            string OwnershipCode = this.InputOwnershipCode.Get(executionContext);
            Money Revenue = this.InputRevenue.Get(executionContext);
            decimal dRev = 0;
            if (!Money.Equals(Revenue, null))
                dRev = Revenue.Value;
            Money BudgetAmount = this.InputBudgetAmount.Get(executionContext);
            decimal dBud = 0;
            if (!Money.Equals(BudgetAmount, null))
                dBud = BudgetAmount.Value;
            bool decisionmaker = this.InputDecisionmaker.Get(executionContext);

            // 営業案件の確度予測を実行
            Value mlresults = InvokeRequestResponseService(webapikey, uri, OwnershipCode, dRev, dBud, decisionmaker);

            // 予測結果を出力情報に設定
            string MLResult = ((string[])(mlresults.Values.GetValue(0)))[0];
            int MLprobability = 0;
            switch (MLResult)
            {
                case "受注":
                    string t = ((string[])(mlresults.Values.GetValue(0)))[1];
                    if (t != null) MLprobability = (int)(double.Parse(t) * 100);
                    break;
                case "失注":
                    break;
                default:
                    break;
            };
            this.OutMLprobability.Set(executionContext, MLprobability);
        }

        // 入力情報: AzureML WebAPI Key
        [RequiredArgument]
        [Input("InputApikey")]
        public InArgument<string> InputApikey { get; set; }

        // 入力情報: AzureML WebURI
        [RequiredArgument]
        [Input("InputUri")]
        public InArgument<string> InputUri { get; set; }

        // 入力情報: 企業形態
        [RequiredArgument]
        [Input("InputOwnershipCode")]
        public InArgument<string> InputOwnershipCode { get; set; }

        // 入力情報: 売上高
        [RequiredArgument]
        [Input("InputRevenue")]
        [Default("0")]
        public InArgument<Money> InputRevenue { get; set; }

        // 入力情報: 予算金額
        [RequiredArgument]
        [Input("InputBudgetAmount")]
        [Default("0")]
        public InArgument<Money> InputBudgetAmount { get; set; }

        // 入力情報: 決済者の有無
        [RequiredArgument]
        [Input("InputDecisionmaker")]
        [Default("True")]
        public InArgument<bool> InputDecisionmaker { get; set; }

        // 出力情報: 予測確度
        [Output("OutMLprobability")]
        public OutArgument<int> OutMLprobability { get; set; }

        static Value InvokeRequestResponseService(string webapikey, string uri, string ownershipcode, decimal revenue, decimal budgetamount, bool decisionmaker)
        {
            Value results = new Value();
            using (var client = new HttpClient())
            {
                string apiKey = webapikey;
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);

                client.BaseAddress = new Uri(uri);

                StringBuilder stb = new StringBuilder();
                stb.Append("{'Inputs':{'input1':{'ColumnNames':['企業形態 (取引先企業) (取引先企業 )','売上高 (取引先企業) (取引先企業 )','予算金額','決定者の有無'],'Values':[[");
                stb.Append(string.Format("'{0}',", ownershipcode));
                stb.Append(string.Format("'{0}',", decimal.ToInt32(revenue).ToString()));
                stb.Append(string.Format("'{0}',", decimal.ToInt32(budgetamount).ToString()));
                stb.Append(string.Format("'{0}',", decisionmaker ? "有" : "無"));
                stb.Append("],[");
                stb.Append(string.Format("'{0}',", ownershipcode));
                stb.Append(string.Format("'{0}',", decimal.ToInt32(revenue).ToString()));
                stb.Append(string.Format("'{0}',", decimal.ToInt32(budgetamount).ToString()));
                stb.Append(string.Format("'{0}',", decisionmaker ? "有" : "無"));
                stb.Append("]]}},'GlobalParameters':{}}");

                string query = stb.ToString();

                HttpResponseMessage response = client.PostAsync("", new StringContent(query, Encoding.UTF8, "application/json")).Result;

                if (response.IsSuccessStatusCode)
                {
                    DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(Rootobject));
                    var stream = response.Content.ReadAsStreamAsync().Result;                    
                    Rootobject p2 = (Rootobject)ser.ReadObject(stream);
                    results = p2.Results.output1.value;
                }
                return results;

            }
        }
    }
}