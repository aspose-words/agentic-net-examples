using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    public class Program
    {
        public static void Main()
        {
            // Ensure the output folder exists.
            const string outputFolder = "Output";
            System.IO.Directory.CreateDirectory(outputFolder);

            // 1. Create a template document with a LINQ Reporting tag that references a missing member.
            var templatePath = System.IO.Path.Combine(outputFolder, "Template.docx");
            var builder = new DocumentBuilder();
            // The tag <<[MissingObject.Name]>> refers to a member that does not exist in the data source.
            builder.Writeln("<<[MissingObject.Name]>>");
            builder.Document.Save(templatePath);

            // 2. Load the template document for reporting.
            var doc = new Document(templatePath);

            // 3. Prepare a data source that does not contain the required member.
            // Using an empty DataSet ensures that "MissingObject" cannot be resolved.
            var dataSource = new DataSet();

            // 4. Create the ReportingEngine without the AllowMissingMembers option.
            var engine = new ReportingEngine();
            // Do NOT set engine.Options = ReportBuildOptions.AllowMissingMembers;
            // The default options (ReportBuildOptions.None) will cause an exception for missing members.

            try
            {
                // 5. Build the report. This should throw because the template references a missing member.
                engine.BuildReport(doc, dataSource, "");
                // If no exception is thrown, indicate unexpected success.
                Console.WriteLine("Report built successfully (unexpected).");
            }
            catch (Exception ex)
            {
                // 6. Expected path: an exception is thrown due to the missing member.
                Console.WriteLine("Exception caught as expected:");
                Console.WriteLine(ex.Message);
            }

            // 7. Save the (potentially unchanged) document to verify the process completed.
            var resultPath = System.IO.Path.Combine(outputFolder, "Result.docx");
            doc.Save(resultPath);
        }
    }
}
