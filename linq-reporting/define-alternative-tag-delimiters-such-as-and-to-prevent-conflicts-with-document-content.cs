using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create an output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Build the template document using the default LINQ Reporting delimiters << >>.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // LINQ Reporting tags.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("<<[item.Index]>> - <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        template.Save(templatePath);

        // 2. Load the template for reporting.
        Document doc = new Document(templatePath);

        // 3. Prepare sample data.
        ReportModel model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Index = 1, Name = "Apple" },
                new Item { Index = 2, Name = "Banana" },
                new Item { Index = 3, Name = "Cherry" }
            }
        };

        // 4. Configure the ReportingEngine.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };

        // 5. Build the report. The root object name is "model" to match the tags.
        engine.BuildReport(doc, model, "model");

        // 6. Save the generated report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(reportPath);
    }
}

// Public data model classes.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
}
