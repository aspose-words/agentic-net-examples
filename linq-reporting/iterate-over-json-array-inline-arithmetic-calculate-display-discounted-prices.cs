using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsJsonReport
{
    class Program
    {
        static void Main()
        {
            // Path to the Word template that contains the reporting placeholders.
            // Example placeholder row in the template:
            // {{Name}} {{Price}} {{Discount}} {{Price * (1 - Discount/100)}}
            string templatePath = @"C:\Templates\DiscountReport.docx";

            // Path to the JSON file. The JSON should be an array of objects, e.g.:
            // [
            //   { "Name": "Product A", "Price": 100.0, "Discount": 10 },
            //   { "Name": "Product B", "Price": 250.0, "Discount": 15 }
            // ]
            string jsonPath = @"C:\Data\Products.json";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Create a JSON data source from the file.
            JsonDataSource jsonData = new JsonDataSource(jsonPath);

            // Build the report. The root element name ("products") must match the
            // name used in the template syntax (e.g. {{#products}} ... {{/products}}).
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, jsonData, "products");

            // Save the generated document.
            string outputPath = @"C:\Output\DiscountReport.docx";
            doc.Save(outputPath);
        }
    }
}
