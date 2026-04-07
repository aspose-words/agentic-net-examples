using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingJson
{
    class Program
    {
        static void Main()
        {
            // Register code page provider (required for some encodings).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // -----------------------------------------------------------------
            // 1. Create sample JSON data file (array of objects).
            // -----------------------------------------------------------------
            string jsonPath = "people.json";
            string jsonContent = @"[
  { ""Name"": ""John Doe"", ""Age"": 30 },
  { ""Name"": ""Jane Smith"", ""Age"": 25 },
  { ""Name"": ""Bob Johnson"", ""Age"": 40 }
]";
            File.WriteAllText(jsonPath, jsonContent);

            // -----------------------------------------------------------------
            // 2. Build a Word template that uses LINQ Reporting tags.
            // -----------------------------------------------------------------
            string templatePath = "template.docx";
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("People Report");
            builder.Writeln("==============");
            // The foreach tag iterates over the JSON data source named 'jsonData'.
            builder.Writeln("<<foreach [in jsonData]>>");
            // Inside the loop we output the fields of each JSON object.
            builder.Writeln("Name: <<[Name]>>, Age: <<[Age]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 3. Load the template and bind the JSON data source.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);
            JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

            // Build the report. The third argument is the name used in the template tags.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, jsonDataSource, "jsonData");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            string outputPath = "output.docx";
            doc.Save(outputPath);

            // The program finishes here without waiting for user input.
        }
    }
}
