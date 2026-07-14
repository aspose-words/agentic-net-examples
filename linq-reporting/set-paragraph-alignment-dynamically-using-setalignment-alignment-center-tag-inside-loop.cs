using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ParagraphItem
{
    public string Text { get; set; } = "";
    public bool Center { get; set; }
}

public class ReportModel
{
    public List<ParagraphItem> Paragraphs { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Paragraphs = new List<ParagraphItem>
            {
                new ParagraphItem { Text = "Left aligned paragraph.", Center = false },
                new ParagraphItem { Text = "Center aligned paragraph.", Center = true },
                new ParagraphItem { Text = "Another left aligned paragraph.", Center = false }
            }
        };

        // Create a template document with a foreach loop.
        var template = new Document();
        var builder = new DocumentBuilder(template);
        builder.Writeln("<<foreach [p in Paragraphs]>>");
        builder.Writeln("<<[p.Text]>>");
        builder.Writeln("<</foreach>>");

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Apply paragraph alignment based on the data source.
        // The generated paragraphs appear in the same order as the source collection.
        var paragraphs = template.FirstSection.Body.Paragraphs;
        for (int i = 0; i < model.Paragraphs.Count && i < paragraphs.Count; i++)
        {
            if (model.Paragraphs[i].Center)
                paragraphs[i].ParagraphFormat.Alignment = ParagraphAlignment.Center;
            else
                paragraphs[i].ParagraphFormat.Alignment = ParagraphAlignment.Left;
        }

        // Save the resulting document.
        template.Save("DynamicAlignmentReport.docx");
    }
}
