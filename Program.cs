using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Net;
using System.Data;
using System.Linq;

// Microsoft Dynamics CRM namespace(s)
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk.Query;

namespace CADAC
{
    /// <summary>
    /// A sandboxed plug-in that can access network (Web) resources.
    /// </summary>
    /// <remarks>Register this plug-in in the sandbox. You can provide an unsecure string
    /// during registration that specifies the Web address (URI) to access from the plug-in.
    /// </remarks>
    public sealed class CADAC : IPlugin
    {

        public void Execute(IServiceProvider serviceProvider)
        {


            ITracingService tracingService =
               (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var _service = serviceFactory.CreateOrganizationService(context.UserId);
            OrganizationServiceContext ctx = new OrganizationServiceContext(_service);


            // The InputParameters collection contains all the data passed in the message request.
            if (context.InputParameters.Contains("Target") &&
                context.InputParameters["Target"] is Entity)
            {
                // Obtain the target entity from the input parameters.
                Entity entity = (Entity)context.InputParameters["Target"];

                // Verify that the target entity represents an inquiry record.
                // If not, this plug-in was not registered correctly.
                if (entity.LogicalName != "annotation")
                    return;

                try
                {



                    Entity fundingApplication = _service.Retrieve("ca_fundingapplication", entity.GetAttributeValue<EntityReference>("objectid").Id, new ColumnSet(true));
                    string documentBody = entity.GetAttributeValue<string>("documentbody");
                    byte[] fileContent = Convert.FromBase64String(documentBody);
                    var dt = ConvertCSVtoDataTable(fileContent.ToString());

                    foreach (DataRow drow in dt.Rows)
                    {
                        string value = drow[1].ToString();
                        Console.WriteLine(value);
                        Entity FundingApplicationLineItem = new Entity("ca_fundingapplicationlineitem");
                        FundingApplicationLineItem["ca_name"] = value;
                        FundingApplicationLineItem["ca_amount"] = Double.Parse(value);
                        _service.Create(FundingApplicationLineItem);
                    }
                }
                catch (WebException exception)
                {
                    throw new InvalidPluginExecutionException("An error has occurred, {0}" + exception.Message);
                }

            }

        }

        public static DataTable ConvertCSVtoDataTable(string strFilePath)
        {
            DataTable dt = new DataTable();
            using (StreamReader sr = new StreamReader(strFilePath))
            {
                string[] headers = sr.ReadLine().Split(',');
                foreach (string header in headers)
                {
                    dt.Columns.Add(header);
                }
                while (!sr.EndOfStream)
                {
                    string[] rows = sr.ReadLine().Split(',');
                    DataRow dr = dt.NewRow();
                    for (int i = 0; i < headers.Length; i++)
                    {
                        dr[i] = rows[i];
                    }
                    dt.Rows.Add(dr);
                }

            }


            return dt;
        }
        public static DataTable ReadCSV(String filename)
        {
            var csvData = new DataTable();
            StreamReader csvFile = null;
            try
            {
                csvFile = new StreamReader(filename);

                // Parse header
                var headerLine = csvFile.ReadLine();
                var columns = ParseCSVLine(headerLine);
                columns.ForEach(c => csvData.Columns.Add(c, typeof(String)));

                var line = "";
                while ((line = csvFile.ReadLine()) != null)
                {
                    if (line == "") // Skip empty line
                        continue;
                    csvData.Rows.Add(
                        ParseCSVLine(line) // Parse CSV Line
                            .OfType<Object>() // Convert it to Object List
                            .ToArray()   // Convert it to Object Array, so that it can be added to DataTable
                    ); // Add Csv Record to Data Table
                }
            }
            finally
            {
                if (csvFile != null)
                    csvFile.Close();
            }

            return csvData;
        }

        private static List<String> ParseCSVLine(String line)
        {
            var quoteStarted = false;
            var values = new List<String>();
            var marker = 0;
            var currPos = 0;
            var prevChar = '\0';

            foreach (Char currChar in line)
            {
                if (currChar == ',' && !quoteStarted)
                {
                    AddValue(line, marker, currPos - marker, values);
                    marker = currPos + 1;
                    quoteStarted = false;
                }
                else if (currChar == '\"')
                    quoteStarted = (prevChar == '\"' && !quoteStarted)
                        ? true
                        : !quoteStarted;
                currPos++;
                prevChar = currChar;
            }
            AddValue(line, marker, currPos - marker, values);
            return values;
        }

        private static void AddValue(String line, Int32 start, Int32 count, List<String> values)
        {
            var val = line.Substring(start, count);
            if (val == "")
                values.Add("");
            else if (val[0] == '\"' && val[val.Length - 1] == '\"')
                values.Add(val.Trim('\"'));
            else
                values.Add(val.Trim());
        }


    }



}
