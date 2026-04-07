using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a blank document and add a paragraph with a <<doc>> tag that references a document.
        // The referenced document may be missing; the engine will treat the missing member as null and skip it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Report start");
        builder.Writeln("<<doc [src.Document]>>"); // Optional include.
        builder.Writeln("Report end");

        // Prepare the data source. If the file exists we load it, otherwise we leave the property null.
        var src = new IncludeSource();
        const string missingPath = "MissingFile.docx";
        if (File.Exists(missingPath))
            src.Document = new Document(missingPath);
        else
            src.Document = null; // Missing file – treated as optional.

        // Configure the reporting engine to ignore missing members (including the null Document).
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.AllowMissingMembers;
        engine.MissingMemberMessage = string.Empty; // No placeholder for missing members.

        // Build the report. The missing include file will be skipped without throwing an exception.
        engine.BuildReport(doc, src, "src");

        // Save the generated document.
        doc.Save("ReportOutput.docx");
    }

    // Wrapper class used as the data source for the report.
    public class IncludeSource
    {
        // When null, the <<doc>> tag is ignored.
        public Document? Document { get; set; }
    }
}
