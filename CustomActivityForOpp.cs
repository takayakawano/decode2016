// =====================================================================
//  This file is part of the Microsoft Dynamics CRM SDK code samples.
//
//  Copyright (C) Microsoft Corporation.  All rights reserved.
//
//  This source code is intended only as a supplement to Microsoft
//  Development Tools and/or on-line documentation.  See these other
//  materials for detailed information regarding Microsoft code samples.
//
//  THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
//  KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
//  IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
//  PARTICULAR PURPOSE.
// =====================================================================

//<snippetCustomActivity>
using System;
using System.Activities;

// These namespaces are found in the Microsoft.Xrm.Sdk.dll assembly
// located in the SDK\bin folder of the SDK download.
using Microsoft.Xrm.Sdk;

// These namespaces are found in the Microsoft.Xrm.Sdk.Workflow.dll assembly
// located in the SDK\bin folder of the SDK download.
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
    /// <summary>
    /// Creates a task with a subject equal to the ID of the input entity.
    /// Input arguments:
    ///   "Input Entity". Type: EntityReference. Is the account entity.
    /// Output argument:
    ///   "Task Created". Type: EntityReference. Is the task created.
    /// </summary>
    public sealed partial class CustomActivityForOpp : CodeActivity
    {

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
        /// Creates a task with a subject equal to the ID of the input EntityReference
        /// </summary>
        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();


            // Input
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

            // Do ML Prediction
            Value mlresults = InvokeRequestResponseService(webapikey, uri, OwnershipCode, dRev, dBud, decisionmaker);

            // Set output paramaters
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

        // Define Input/Output Arguments
        [RequiredArgument]
        [Input("InputApikey")]
        public InArgument<string> InputApikey { get; set; }

        [RequiredArgument]
        [Input("InputUri")]
        public InArgument<string> InputUri { get; set; }

        [RequiredArgument]
        [Input("InputOwnershipCode")]
        public InArgument<string> InputOwnershipCode { get; set; }

        [RequiredArgument]
        [Input("InputRevenue")]
        [Default("0")]
        public InArgument<Money> InputRevenue { get; set; }

        [RequiredArgument]
        [Input("InputBudgetAmount")]
        [Default("0")]
        public InArgument<Money> InputBudgetAmount { get; set; }

        [RequiredArgument]
        [Input("InputDecisionmaker")]
        [Default("True")]
        public InArgument<bool> InputDecisionmaker { get; set; }

        [Output("OutMLprobability")]
        public OutArgument<int> OutMLprobability { get; set; }

        static Value InvokeRequestResponseService(string webapikey, string uri, string ownershipcode, decimal revenue, decimal budgetamount, bool decisionmaker)
        {
            Value results = new Value();
            using (var client = new HttpClient())
            {

                string apiKey = webapikey; // Replace this with the API key for the web service
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);

                client.BaseAddress = new Uri(uri);

                StringBuilder stb = new StringBuilder();
                stb.Append("{'Inputs':{'input1':{'ColumnNames':['企業形態 (取引先企業) (取引先企業 )','売上高 (取引先企業) (取引先企業 )','予算金額','決定者の有無'],'Values':[[");
                stb.Append(string.Format("'{0}',", ownershipcode));
                stb.Append(string.Format("'{0}',", decimal.ToInt32(revenue).ToString()));//"100000000"));
                stb.Append(string.Format("'{0}',", decimal.ToInt32(budgetamount).ToString()));// "1000000")); //budgetamount.Value.ToString()));
                stb.Append(string.Format("'{0}',", decisionmaker ? "有" : "無"));
                stb.Append("],[");
                stb.Append(string.Format("'{0}',", ownershipcode));
                stb.Append(string.Format("'{0}',", decimal.ToInt32(revenue).ToString()));//"100000000"));
                stb.Append(string.Format("'{0}',", decimal.ToInt32(budgetamount).ToString())); //budgetamount.Value.ToString()));
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
//</snippetCustomActivity>