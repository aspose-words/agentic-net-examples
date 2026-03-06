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

        // Initialize a DocumentBuilder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a template that references a member which does not exist in the data source.
        // The syntax <<[...]>>
        // will be processed by ReportingEngine.
        builder.Writeln("<<[missingObject.First().id]>>");
        builder.Writeln("<<foreach [in missingObject]>><<[id]>><</foreach>>");

        // Configure the ReportingEngine:
        // - AllowMissingMembers enables the engine to continue when a member is missing.
        // - MissingMemberMessage defines the text that will be inserted instead of the missing value.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers,
            MissingMemberMessage = "Missed"
        };

        // Build the report using an empty DataSet (no "missingObject" defined).
        // The engine will replace the missing references with the message defined above.
        engine.BuildReport(builder.Document, new DataSet(), string.Empty);

        // Save the resulting document as a Markdown file.
        // MarkdownSaveOptions allows us to control how links are exported.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            // Example: export all links as reference blocks.
            LinkExportMode = MarkdownLinkExportMode.Reference
        };

        doc.Save("Report.md", saveOptions);
    }
}
