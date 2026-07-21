using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model used by the LINQ Reporting engine.
    public class OrderModel
    {
        // Quantity is stored as double to demonstrate explicit casting to int.
        public double Quantity { get; set; } = 0;

        // Method that expects an integer parameter.
        public string GetQuantityMessage(int qty)
        {
            return $"Quantity (int) = {qty}";
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Step 1: Create a template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert a LINQ Reporting tag that calls GetQuantityMessage,
            // explicitly casting the double Quantity to int.
            builder.Writeln("<<[model.GetQuantityMessage((int)model.Quantity)]>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // Step 2: Load the template for reporting.
            Document reportDoc = new Document(templatePath);

            // Prepare the data source.
            OrderModel model = new OrderModel
            {
                Quantity = 12.7 // Example value that will be cast to int (12).
            };

            // Step 3: Build the report using ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Step 4: Save the generated report.
            const string outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
