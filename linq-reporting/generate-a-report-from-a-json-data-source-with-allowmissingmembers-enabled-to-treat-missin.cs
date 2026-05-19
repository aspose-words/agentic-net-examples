using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Reporting;
using System.Text;

namespace AsposeWordsLinqReporting
{
    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some encodings).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for the template, JSON data and the final report.
            const string templatePath = "template.docx";
            const string jsonPath = "data.json";
            const string reportPath = "report.docx";

            // -----------------------------------------------------------------
            // 1. Create a simple JSON file with some missing fields.
            // -----------------------------------------------------------------
            string jsonContent = @"
[
    { ""Name"": ""John Doe"", ""Age"": 30 },
    { ""Name"": ""Alice Smith"", ""Age"": 25, ""Email"": ""alice@example.com"" }
]";
            File.WriteAllText(jsonPath, jsonContent);

            // -----------------------------------------------------------------
            // 2. Build a Word template programmatically and save it.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("People Report");
            builder.Writeln("<<foreach [p in persons]>>");
            builder.Writeln("Name : <<[p.Name]>>");
            builder.Writeln("Age  : <<[p.Age]>>");
            // Email may be missing; AllowMissingMembers will treat it as null.
            builder.Writeln("Email: <<[p.Email]>>");
            builder.Writeln("<</foreach>>");

            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 3. Load the template and prepare the JSON data source.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);
            JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

            // -----------------------------------------------------------------
            // 4. Configure the reporting engine to allow missing members.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.AllowMissingMembers;
            engine.MissingMemberMessage = "N/A";

            // Build the report. The root name "persons" matches the JSON array.
            engine.BuildReport(reportDoc, jsonDataSource, "persons");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            reportDoc.Save(reportPath);
        }
    }
}
