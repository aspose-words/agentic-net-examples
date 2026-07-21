using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Required for Table type

namespace AsposeWordsLinqReportingExample
{
    // Data model for a line item.
    public class LineItem
    {
        public string Description { get; set; } = string.Empty;
        public int Quantity { get; set; }
        public decimal UnitPrice { get; set; }
    }

    // Wrapper model containing the collection used in the foreach band.
    public class ReportModel
    {
        public List<LineItem> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words (required in .NET Core).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Create a simple template document with a foreach band.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Begin foreach band over Items collection.
            builder.Writeln("<<foreach [item in Items]>>");

            // Create a table to display each line item.
            Table table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Writeln("Description");
            builder.InsertCell();
            builder.Writeln("Quantity");
            builder.InsertCell();
            builder.Writeln("Unit Price");
            builder.InsertCell();
            builder.Writeln("Line Total");
            builder.EndRow();

            // Data row – use expression tags to output properties and calculate total.
            builder.InsertCell();
            builder.Writeln("<<[item.Description]>>");
            builder.InsertCell();
            builder.Writeln("<<[item.Quantity]>>");
            builder.InsertCell();
            builder.Writeln("<<[item.UnitPrice]>>");
            builder.InsertCell();
            // Calculate line total directly in the expression tag.
            builder.Writeln("<<[item.Quantity * item.UnitPrice]>>");
            builder.EndRow();

            builder.EndTable();

            // End foreach band.
            builder.Writeln("<</foreach>>");

            // Prepare sample data.
            ReportModel model = new ReportModel
            {
                Items = new List<LineItem>
                {
                    new LineItem { Description = "Apple", Quantity = 3, UnitPrice = 1.5m },
                    new LineItem { Description = "Banana", Quantity = 2, UnitPrice = 0.8m },
                    new LineItem { Description = "Cherry", Quantity = 5, UnitPrice = 0.6m }
                }
            };

            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            doc.Save("Report.docx");
        }
    }
}
