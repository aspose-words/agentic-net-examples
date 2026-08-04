using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model with a missing member (MissingInfo) to demonstrate AllowMissingMembers.
    public class Order
    {
        public string CustomerName { get; set; } = "John Doe";
        // Note: No MissingInfo property – it will be treated as null by the engine.
    }

    public class Program
    {
        public static void Main()
        {
            // Create a template document programmatically.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Insert LINQ Reporting tags. The second tag references a missing member.
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("Missing member value: <<[order.MissingInfo]>>");

            // Configure the reporting engine to treat missing members as null.
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.AllowMissingMembers;
            // Optional: customize the message shown for a plain missing member reference.
            engine.MissingMemberMessage = "";

            // Build the report using the data source and the root name "order".
            var order = new Order();
            engine.BuildReport(doc, order, "order");

            // Save the generated report.
            doc.Save("Report.docx");
        }
    }
}
