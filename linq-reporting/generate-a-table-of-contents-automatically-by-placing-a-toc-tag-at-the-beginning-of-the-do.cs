using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Ensure code page support for possible Unicode data.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Step 1: Create the template document programmatically.
        var templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Step 2: Load the template for reporting.
        var doc = new Document(templatePath);

        // Step 3: Prepare a simple data model.
        var model = new ReportModel
        {
            Title = "Sample Report",
            Sections = new List<SectionModel>
            {
                new SectionModel { Heading = "Chapter 1", SubHeading = "Section 1.1" },
                new SectionModel { Heading = "Chapter 2", SubHeading = "Section 2.1" }
            }
        };

        // Step 4: Build the report using LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options
        engine.BuildReport(doc, model, "model");

        // Step 5: Update fields so the TOC reflects the generated headings.
        doc.UpdateFields();

        // Step 6: Save the final document.
        doc.Save("Report.docx");
    }

    private static void CreateTemplate(string filePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert a TOC field that will later be updated.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Insert a title placeholder using LINQ Reporting tag.
        builder.Writeln("<<[model.Title]>>");
        builder.InsertBreak(BreakType.PageBreak);

        // Add sample headings that will appear in the TOC.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("<<[model.Sections[0].Heading]>>");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("<<[model.Sections[0].SubHeading]>>");
        builder.InsertParagraph();

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("<<[model.Sections[1].Heading]>>");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("<<[model.Sections[1].SubHeading]>>");

        // Save the template to disk.
        doc.Save(filePath);
    }
}

// Public data model aligned with the template tags.
public class ReportModel
{
    public string Title { get; set; } = string.Empty;
    public List<SectionModel> Sections { get; set; } = new();
}

public class SectionModel
{
    public string Heading { get; set; } = string.Empty;
    public string SubHeading { get; set; } = string.Empty;
}
