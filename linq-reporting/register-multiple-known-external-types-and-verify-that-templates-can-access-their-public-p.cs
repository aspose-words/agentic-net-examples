using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // External type 1 with a public static property.
    public static class ExternalType1
    {
        public static string Name => "ExternalOne";
    }

    // External type 2 with a public static property.
    public static class ExternalType2
    {
        public static int Value => 42;
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document that references the external types.
            // -----------------------------------------------------------------
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // Insert LINQ Reporting tags that access static properties of the known types.
            builder.Writeln("Type1: <<[ExternalType1.Name]>>");
            builder.Writeln("Type2: <<[ExternalType2.Value]>>");

            // Save the template to disk (required by the lifecycle rules).
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // Load the template back before building the report.
            var loadedTemplate = new Document(templatePath);

            // ---------------------------------------------------------------
            // 2. Configure the ReportingEngine and register the external types.
            // ---------------------------------------------------------------
            var engine = new ReportingEngine();
            engine.KnownTypes.Add(typeof(ExternalType1));
            engine.KnownTypes.Add(typeof(ExternalType2));

            // Build the report. No data source is needed; we pass an empty object.
            // The empty string for the data source name allows direct access to static members.
            bool success = engine.BuildReport(loadedTemplate, new object(), "");

            // ---------------------------------------------------------------
            // 3. Verify the result and save the final document.
            // ---------------------------------------------------------------
            if (success)
            {
                // Extract the plain text of the generated document.
                string resultText = loadedTemplate.GetText();

                // Output the result to the console for verification.
                Console.WriteLine("Report generation succeeded. Document content:");
                Console.WriteLine(resultText);
            }
            else
            {
                Console.WriteLine("Report generation failed.");
            }

            // Save the populated document.
            const string resultPath = "Result.docx";
            loadedTemplate.Save(resultPath);
        }
    }
}
