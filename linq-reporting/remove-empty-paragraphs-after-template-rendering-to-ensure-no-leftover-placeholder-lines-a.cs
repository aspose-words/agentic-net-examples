using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

#nullable enable

public class Program
{
    public static void Main()
    {
        // Prepare sample data model.
        var model = new ReportModel
        {
            Name = "John Doe",
            // OptionalLine is left null to produce an empty paragraph after rendering.
            OptionalLine = null
        };

        // Create a template document programmatically.
        var templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Load the template.
        var doc = new Document(templatePath);

        // Build the report using LINQ Reporting Engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Remove empty paragraphs that may have been left after rendering.
        RemoveEmptyParagraphs(doc);

        // Save the final document.
        var outputPath = "Report.docx";
        doc.Save(outputPath);
    }

    private static void CreateTemplate(string path)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Greeting line with a placeholder.
        builder.Writeln("Hello <<[model.Name]>>!");

        // This line contains only a placeholder that may become empty.
        builder.Writeln("<<[model.OptionalLine]>>");

        // Save the template.
        doc.Save(path);
    }

    private static void RemoveEmptyParagraphs(Document doc)
    {
        // Collect paragraphs to remove to avoid modifying the collection while iterating.
        var paragraphsToRemove = new List<Paragraph>();

        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // Get the paragraph text without the end-of-paragraph marker.
            var text = para.GetText().Replace("\a", string.Empty).Trim();

            if (string.IsNullOrEmpty(text))
                paragraphsToRemove.Add(para);
        }

        foreach (var para in paragraphsToRemove)
            para.Remove();
    }
}

// Data model used by the report.
public class ReportModel
{
    public string Name { get; set; } = string.Empty;
    public string? OptionalLine { get; set; }
}
