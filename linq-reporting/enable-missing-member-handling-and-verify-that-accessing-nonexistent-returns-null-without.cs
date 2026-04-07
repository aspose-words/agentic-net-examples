using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    public class Program
    {
        public static void Main()
        {
            // Create a template that contains a tag referencing a missing member.
            const string templatePath = "Template.docx";
            CreateTemplate(templatePath);

            // Load the template document.
            var doc = new Document(templatePath);

            // Data model without the referenced member.
            var model = new Model();

            // Configure the reporting engine to treat missing members as null.
            var engine = new ReportingEngine
            {
                Options = ReportBuildOptions.AllowMissingMembers,
                MissingMemberMessage = "" // Optional: keep the placeholder empty.
            };

            // Build the report using the model as the root object named "model".
            engine.BuildReport(doc, model, "model");

            // Verify that the missing member tag produced no output (null/empty).
            var result = doc.GetText().Trim();

            Console.WriteLine("Resulting document text:");
            Console.WriteLine(result);
            Console.WriteLine("Missing member handled without error: " + (result == string.Empty));
        }

        private static void CreateTemplate(string filePath)
        {
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Tag that tries to access a non‑existent property.
            builder.Writeln("<<[model.Nonexistent]>>");

            doc.Save(filePath);
        }

        // Simple model class with no properties.
        public class Model
        {
        }
    }
}
