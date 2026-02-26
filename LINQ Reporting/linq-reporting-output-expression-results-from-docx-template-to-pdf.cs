using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace LinqReportingToPdf
{
    // Simple data source class used by the reporting engine.
    public class ReportData
    {
        // Example numeric value that will be formatted in the template.
        public decimal Amount { get; set; }

        // Example string value.
        public string CustomerName { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a blank Word document that will serve as the template.
            // -----------------------------------------------------------------
            Document template = new Document();                     // create
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert template tags. The tags are enclosed in << >> and reference
            // members of the data source object (named "data" below).
            // The ":dollarText" format converts a numeric value to its textual
            // representation (e.g., 123.45 -> "One Hundred Twenty‑Three Dollars").
            builder.Writeln("Customer: <<[data.CustomerName]>>");
            builder.Writeln("Amount: <<[data.Amount]:dollarText>>");

            // -----------------------------------------------------------------
            // 2. Prepare the data source instance.
            // -----------------------------------------------------------------
            ReportData data = new ReportData
            {
                CustomerName = "Acme Corp.",
                Amount = 1234.56m
            };

            // -----------------------------------------------------------------
            // 3. Build the report using the LINQ Reporting engine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();

            // The second parameter is the data source object.
            // The third parameter ("data") is the name used inside the template.
            engine.BuildReport(template, data, "data");

            // -----------------------------------------------------------------
            // 4. Save the populated document as PDF.
            // -----------------------------------------------------------------
            // Using the overload that specifies the format explicitly.
            template.Save("ReportOutput.pdf", SaveFormat.Pdf);      // save
        }
    }
}
