using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using ExcelDataReader;
using QuickGraph;
using QuickGraph.Algorithms;
using QuickGraph.Graphviz;
using QuickGraph.Graphviz.Dot;

namespace OpFinder
{
    public class Program
    {
        private const string EXCEL_TAB_NAME = "Report 4";
        private const string EXCEL_FILE_PATH = "data2.xlsx";
        private const string OUT_PATH = "out.csv";

        public static void Main(string[] args)
        {
            Console.WriteLine("Reading Excel..");
            // register encodings because excel is formatted as fook
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);


            using var stream = File.Open(EXCEL_FILE_PATH, FileMode.Open, FileAccess.Read);
            using var reader = ExcelReaderFactory.CreateReader(stream);
            var excelDataset = reader.AsDataSet();
            Console.WriteLine($"read tables:[{excelDataset.Tables.Count}]");

            
            // GRAPH of client, eng and invoices
            var graph = new AdjacencyGraph<string, Edge<string>>();


            // ClientId -> EngId -> InvoiceIds
            var invoiceDictionary = new Dictionary<string, Dictionary<string, List<string>>>();
            var dataRows = excelDataset.Tables[EXCEL_TAB_NAME].Rows;

            // InvoiceId -> EngIds
            var engsPerInvoice = new Dictionary<string, List<string>>();

            Console.WriteLine("Reading data into dictionary");
            for (var i = 1; i < dataRows.Count; i++)
            {
                var invoiceId = dataRows[i][0].ToString().Trim();
                var engId = dataRows[i][1].ToString().Trim();
                var clientId = dataRows[i][2].ToString().Trim();

                if (!invoiceDictionary.ContainsKey(clientId)) { invoiceDictionary[clientId] = new Dictionary<string, List<string>>(); }
                if (!invoiceDictionary[clientId].ContainsKey(engId)) { invoiceDictionary[clientId][engId] = new List<string>(); }
                invoiceDictionary[clientId][engId].Add(invoiceId);
                if (!engsPerInvoice.ContainsKey(invoiceId)) { engsPerInvoice[invoiceId] = new List<string>(); }

                engsPerInvoice[invoiceId].Add(engId);

                
                graph.AddVerticesAndEdge(new Edge<string>($"{clientId}-{invoiceId}", $"{clientId}-{engId}"));
                graph.AddVerticesAndEdge(new Edge<string>($"{clientId}-{engId}", $"{clientId}-{invoiceId}"));
            }

            var componentCount = graph.StronglyConnectedComponents(out var components);
            Console.WriteLine($"calculated OP count:[{componentCount}]");


            var results = new List<ResultRow>();
            foreach (var c in components)
            {
                var clientId = c.Key.Split('-')[0];
                var engagementOrInvoiceId = c.Key.Split('-')[1];
                if (engagementOrInvoiceId.StartsWith("IL")){continue;}
                var resultItem = new ResultRow()
                {
                    ClientId = clientId,
                    TempOpId = $"OP-{c.Value}-{clientId}",
                    EngagementOrInvoiceId = engagementOrInvoiceId
                };
                results.Add(resultItem);
            }
            
            WriteOutput(results);
            
        }

        private static void WriteOutput(List<ResultRow> results)
        {
            using var w = new StreamWriter(OUT_PATH);
            foreach (var r in results)
            {
                var line = $"{r.ClientId},{r.TempOpId},{r.EngagementOrInvoiceId}";
                w.WriteLine(line);
                w.Flush();
            }
            w.Flush();
            w.Close();
        }
    }
}