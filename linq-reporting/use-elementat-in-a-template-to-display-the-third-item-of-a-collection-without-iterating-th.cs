using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary template and the final report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create a simple Word template that contains a LINQ Reporting tag.
        //    The tag uses ElementAt(2) to fetch the third item (zero‑based index).
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Third item: <<[model.Items.ElementAt(2).Name]>>");
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template back (required by the reporting workflow).
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the data model with a collection of items.
        // -----------------------------------------------------------------
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "First" },
                new Item { Name = "Second" },
                new Item { Name = "Third" },   // This will be displayed.
                new Item { Name = "Fourth" }
            }
        };

        // -----------------------------------------------------------------
        // 4. Build the report using the ReportingEngine.
        //    The root object name in the template is "model".
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        loadedTemplate.Save(reportPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes required by the template.
// ---------------------------------------------------------------------
public class ReportModel
{
    // Initialise to avoid nullable warnings.
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    // Initialise to avoid nullable warnings.
    public string Name { get; set; } = string.Empty;
}
