using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a template document that contains a LINQ Reporting placeholder.
        Document template = new Document();                     // create blank document
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Report Header");                      // any static content
        builder.Writeln("<<doc [src.Document]>>");             // placeholder for the RTF document
        builder.Writeln("Report Footer");                      // any static content

        // Load the source document that is in RTF format.
        Document rtfSource = new Document("Source.rtf");        // load existing RTF file

        // Use the ReportingEngine to replace the placeholder with the loaded RTF document.
        ReportingEngine engine = new ReportingEngine();        // create reporting engine
        // The data source array contains the document to insert,
        // and the corresponding name ("src") matches the placeholder prefix.
        engine.BuildReport(template, new object[] { rtfSource }, new string[] { "src" });

        // Save the final merged document.
        template.Save("MergedResult.docx");                    // save to DOCX (format inferred from extension)
    }
}
