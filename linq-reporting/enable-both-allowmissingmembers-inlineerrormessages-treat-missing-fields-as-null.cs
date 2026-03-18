using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

class ReportingEngineExample
{
    static void Main()
    {
        // Create a template document that contains:
        // 1. A reference to a missing member (will be treated as null).
        // 2. An invalid template syntax (will be inlined as an error message).
        DocumentBuilder builder = new DocumentBuilder();
        builder.Writeln("<<[missingObject.Id]>>");          // missing member
        builder.Writeln("<<[invalid syntax>>");            // syntax error (unclosed tag)

        // Configure the ReportingEngine:
        // - AllowMissingMembers: missing members become null literals.
        // - InlineErrorMessages: syntax errors are written into the output instead of throwing.
        // - MissingMemberMessage set to empty string so missing members render as empty (null).
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers | ReportBuildOptions.InlineErrorMessages,
            MissingMemberMessage = ""
        };

        // Build the report using an empty data source (DataSet with no tables).
        // The engine will process the template according to the options above.
        bool parsingSuccessful = engine.BuildReport(builder.Document, new DataSet());

        // The returned flag is false when there are syntax errors, but the document now contains
        // the inline error messages thanks to InlineErrorMessages option.
        Console.WriteLine($"Parsing successful: {parsingSuccessful}");

        // Save the resulting document.
        builder.Document.Save("ReportingEngine_Output.docx");
    }
}
