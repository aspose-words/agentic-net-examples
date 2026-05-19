using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model with only a Name property.
    public class Customer
    {
        public string Name { get; set; } = "John Doe";
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document and add LINQ Reporting tags.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Tag that references an existing member.
            builder.Writeln("Customer Name: <<[customer.Name]>>");
            // Tag that references a missing member – will be treated as null.
            builder.Writeln("Missing Field: <<[customer.MissingField]>>");

            // Prepare the data source.
            Customer customer = new Customer();

            // Configure the reporting engine to allow missing members.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.AllowMissingMembers;
            // Optional: customize the message shown for a plain missing member reference.
            engine.MissingMemberMessage = string.Empty;

            // Build the report. The root object name must match the tag prefix.
            engine.BuildReport(doc, customer, "customer");

            // Save the resulting document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report_AllowMissingMembers.docx");
            doc.Save(outputPath);

            Console.WriteLine($"Report generated: {outputPath}");
        }
    }
}
