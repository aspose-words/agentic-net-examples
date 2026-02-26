using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // RTF fragment that we want to insert into the report.
        // This string is a valid RTF snippet.
        string rtfContent = @"{\rtf1\ansi\deff0 {\fonttbl{\f0\fswiss Arial;}} \f0\fs24 This is \b bold\b0  and \i italic\i0  text.}";

        // Insert a LINQ Reporting placeholder.
        // The placeholder uses contextual object member access (data.RtfContent) and the ':rtf' switch
        // tells the ReportingEngine to treat the value as Rich Text Format.
        builder.Writeln("<<[data.RtfContent]:rtf>>");

        // Create an anonymous object that serves as the data source.
        var dataSource = new { RtfContent = rtfContent };

        // Build the report. This overload allows the template to reference the data source object itself.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "data");

        // Save the populated document.
        doc.Save("ReportWithRtf.docx");
    }
}
