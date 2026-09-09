using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingRestrictedMembers
{
    // Simple data model used by the report.
    public class Model
    {
        public string Name { get; set; } = "John Doe";
        public string Secret { get; set; } = "TopSecret";
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure the output directory exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Paths for the template and the generated report.
            string templatePath = Path.Combine(outputDir, "template.docx");
            string reportPath = Path.Combine(outputDir, "report.docx");

            // -----------------------------------------------------------------
            // 1. Create a Word template containing LINQ Reporting tags.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Normal property access – allowed.
            builder.Writeln("Name: <<[model.Name]>>");

            // Accessing a member of System.Type – will be restricted.
            builder.Writeln("Type: <<[model.GetType().FullName]>>");

            // Another normal property – allowed.
            builder.Writeln("Secret: <<[model.Secret]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and configure the ReportingEngine.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // Restrict access to System.Type and its members.
            ReportingEngine.SetRestrictedTypes(typeof(System.Type));

            ReportingEngine engine = new ReportingEngine
            {
                // Allow missing members so the engine does not throw an exception
                // when a restricted member is accessed.
                Options = ReportBuildOptions.AllowMissingMembers,
                MissingMemberMessage = "[Restricted]"
            };

            // -----------------------------------------------------------------
            // 3. Build the report using a model instance.
            // -----------------------------------------------------------------
            Model model = new Model();

            // The root object name used in the template tags is "model".
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            doc.Save(reportPath);

            // Inform the user where the files are located.
            Console.WriteLine($"Template saved to: {templatePath}");
            Console.WriteLine($"Report saved to:   {reportPath}");
        }
    }
}
