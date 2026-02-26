using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class InsertRtfWithLinqReporting
{
    static void Main()
    {
        // Path to the source RTF document that will be inserted.
        string rtfPath = @"C:\Docs\SourceDocument.rtf";

        // Path where the final merged document will be saved.
        string outputPath = @"C:\Docs\MergedResult.docx";

        // 1. Load the source RTF document.
        Document sourceDoc = new Document(rtfPath);

        // 2. Create a template document in memory.
        Document template = new Document();                     // create blank document
        DocumentBuilder builder = new DocumentBuilder(template); // create builder for the template

        // Insert a LINQ Reporting placeholder that will be replaced with the source document.
        // The placeholder syntax <<doc [src.Document]>> tells the engine to insert the whole document.
        builder.Writeln("<<doc [src.Document]>>");

        // 3. Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();

        // The data source array contains the document to be inserted.
        // The corresponding name "src" is used in the placeholder.
        engine.BuildReport(template, new object[] { sourceDoc }, new string[] { "src" });

        // 4. Save the merged document.
        template.Save(outputPath);
    }
}
