using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a blank document that will serve as the LINQ Reporting template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // Template: iterate over sections.
        // -----------------------------------------------------------------
        builder.Writeln("<<foreach [sec in Sections]>>");
        // Section title.
        builder.Writeln("<<[sec.Title]>>");
        // Start a numbered list for the items of the current section.
        builder.ListFormat.ApplyNumberDefault();
        // Restart numbering for each new section and iterate over its items.
        builder.Writeln("<<restartNum>><<foreach [it in sec.Items]>><<[it]>><</foreach>>");
        // End of outer foreach.
        builder.Writeln("<</foreach>>");

        // -----------------------------------------------------------------
        // Prepare sample data.
        // -----------------------------------------------------------------
        ReportModel model = new()
        {
            Sections = new List<Section>
            {
                new Section
                {
                    Title = "Fruits",
                    Items = new List<string> { "Apple", "Banana", "Cherry" }
                },
                new Section
                {
                    Title = "Vegetables",
                    Items = new List<string> { "Carrot", "Lettuce", "Tomato" }
                }
            }
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save("RestartNumberingReport.docx");
    }
}

// ---------------------------------------------------------------------
// Data model used by the LINQ Reporting template.
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Section> Sections { get; set; } = new();
}

public class Section
{
    public string Title { get; set; } = string.Empty;
    public List<string> Items { get; set; } = new();
}
