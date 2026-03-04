using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a markdown template that references a member which does not exist in the data source.
        // The placeholder will be processed by the ReportingEngine.
        builder.Writeln("<<[MissingObject.First().Id]>>");

        // Configure the ReportingEngine to allow missing members and provide a custom message.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers,
            MissingMemberMessage = "Missing"
        };

        // Build the report using an empty DataSet (so the referenced member is indeed missing).
        engine.BuildReport(doc, new DataSet(), string.Empty);

        // Set up Markdown save options to export links inline.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            LinkExportMode = MarkdownLinkExportMode.Inline
        };

        // Save the resulting document as a Markdown file.
        doc.Save("Report.md", saveOptions);
    }
}
