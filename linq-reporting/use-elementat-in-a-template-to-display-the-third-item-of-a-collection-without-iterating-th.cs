using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "First",  Index = 1 },
                new Item { Name = "Second", Index = 2 },
                new Item { Name = "Third",  Index = 3 },
                new Item { Name = "Fourth", Index = 4 }
            }
        };

        // -----------------------------------------------------------------
        // 1. Create a template document with a LINQ Reporting tag that uses
        //    ElementAt to fetch the third element (index 2) from the collection.
        // -----------------------------------------------------------------
        string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Third item name: <<[model.Items.ElementAt(2).Name]>>");
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // 3. Save the generated report.
        // -----------------------------------------------------------------
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}

// Root data model referenced in the template as "model".
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Simple item class used in the collection.
public class Item
{
    public string Name { get; set; } = string.Empty;
    public int Index { get; set; }
}
