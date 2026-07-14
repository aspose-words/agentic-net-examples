using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Initialize the collection to avoid nullable warnings.
    public List<string> Tags { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required by Aspose.Words for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for the template and the generated reports.
        string templatePath = "Template.docx";
        string reportWithTagsPath = "Report_WithTags.docx";
        string reportWithoutTagsPath = "Report_WithoutTags.docx";

        // -------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Header.
        builder.Writeln("Tag List:");

        // Conditional block: display the list only if the collection has any items.
        builder.Writeln("<<if [model.Tags.Any()]>>");

        // Loop through each tag.
        builder.Writeln("<<foreach [tag in model.Tags]>>- <<[tag]>>");
        builder.Writeln("<</foreach>>");

        // End of conditional block.
        builder.Writeln("<</if>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template and build a report with tags.
        // -------------------------------------------------
        Document docWithTags = new Document(templatePath);

        ReportModel modelWithTags = new ReportModel
        {
            Tags = new List<string> { "Alpha", "Beta", "Gamma" }
        };

        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(docWithTags, modelWithTags, "model");
        docWithTags.Save(reportWithTagsPath);

        // -------------------------------------------------
        // 3. Build a report where the tag collection is empty.
        // -------------------------------------------------
        Document docWithoutTags = new Document(templatePath);

        ReportModel modelWithoutTags = new ReportModel(); // Tags list is empty.

        engine.BuildReport(docWithoutTags, modelWithoutTags, "model");
        docWithoutTags.Save(reportWithoutTagsPath);
    }
}
