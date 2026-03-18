using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Vba;

// Add the required assembly attribute to enable merging of the ReportingEngine
// with another assembly during deployment.
[assembly: VbaExecutable]

public class Program
{
    public static void Main()
    {
        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, Aspose.Words Reporting Engine!");

        // Initialize the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();

        // Build the report (no data source in this minimal example).
        engine.BuildReport(doc, new object());

        // Save the document.
        doc.Save("ReportingEngineMerged.docx");
    }
}
