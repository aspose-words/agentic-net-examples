using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Use DocumentBuilder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the paragraph style to Heading1 and write the required heading text.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("LINQ Reporting Introduction to LINQ Reporting Engine");

        // Demonstrate the LINQ Reporting Engine – build the report with an empty data source.
        ReportingEngine engine = new ReportingEngine();
        // The BuildReport method returns true if parsing succeeds.
        engine.BuildReport(doc, new object());

        // Save the resulting document to disk.
        doc.Save("LINQReportingHeading.docx");
    }
}
