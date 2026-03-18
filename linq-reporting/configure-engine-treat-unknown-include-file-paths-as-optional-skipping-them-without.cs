// Program.cs
using System;
using System.Data;                     // For DataSet
using Aspose.Words;                     // Core Aspose.Words types
using Aspose.Words.Reporting;           // ReportingEngine and ReportBuildOptions

class Program
{
    static void Main()
    {
        // Create a new empty document and a DocumentBuilder attached to it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an INCLUDE field that points to a non‑existent file.
        // The field syntax is <<include "MissingFile.docx">>
        builder.Writeln("<<include \"MissingFile.docx\">>");

        // Configure the ReportingEngine to treat syntax errors (such as a missing include file)
        // as inline messages instead of throwing exceptions. This effectively skips the missing file.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.InlineErrorMessages
        };

        // Build the report. The engine will process the INCLUDE field, encounter the missing file,
        // and because InlineErrorMessages is enabled, it will skip the field without raising an error.
        engine.BuildReport(doc, new DataSet(), "");

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}
