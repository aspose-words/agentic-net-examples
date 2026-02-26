using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Simple data class to be used as a data source for the report.
    public class ReportData
    {
        public string Title { get; set; }
        public string Description { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Create a new blank Word document.
            Document doc = new Document();

            // Use DocumentBuilder to add a heading to the document.
            DocumentBuilder builder = new DocumentBuilder(doc);
            // Set the paragraph style to Heading 1.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            // Write the heading text.
            builder.Writeln("LINQ Reporting Introduction to LINQ Reporting Engine");

            // Add a placeholder for the report data using LINQ Reporting syntax.
            // The placeholder will be replaced by the ReportingEngine.
            builder.Writeln("<<[data.Title]>>");
            builder.Writeln("<<[data.Description]>>");

            // Prepare the data source.
            ReportData data = new ReportData
            {
                Title = "Welcome to LINQ Reporting",
                Description = "This document demonstrates how to use Aspose.Words LINQ Reporting Engine."
            };

            // Build the report by populating the template with the data source.
            ReportingEngine engine = new ReportingEngine();
            // The data source name "data" matches the placeholder used in the template.
            engine.BuildReport(doc, data, "data");

            // Save the resulting document to a file.
            doc.Save("LinqReportingIntroduction.docx");
        }
    }
}
