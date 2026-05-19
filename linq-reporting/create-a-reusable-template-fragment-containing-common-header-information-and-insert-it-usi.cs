using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Required for code page support (e.g., when using certain data sources).
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // Create a reusable header fragment that contains the report title.
        // -----------------------------------------------------------------
        Document headerFragment = new Document();
        DocumentBuilder headerBuilder = new DocumentBuilder(headerFragment);
        // The title will be filled from the root data source named "model".
        headerBuilder.Writeln("Report Title: <<[model.Title]>>");
        headerFragment.Save("HeaderFragment.docx");

        // ---------------------------------------------------------------
        // Build the main template document.
        // It will include the header fragment and then iterate over Items.
        // ---------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(template);

        // Insert the external header fragment directly into the template.
        // This avoids the need for the <<doc>> tag, which expects a data source.
        Document loadedHeader = new Document("HeaderFragment.docx");
        templateBuilder.InsertDocument(loadedHeader, ImportFormatMode.KeepSourceFormatting);

        // Simple foreach loop to list items.
        templateBuilder.Writeln("<<foreach [item in Items]>>");
        templateBuilder.Writeln("Item <<[item.Index]>>: <<[item.Name]>>");
        templateBuilder.Writeln("<</foreach>>");

        template.Save("Template.docx");

        // ---------------------------------------------------------------
        // Load the template and generate the final report.
        // ---------------------------------------------------------------
        Document report = new Document("Template.docx");

        // Sample data model.
        ReportModel model = new()
        {
            Title = "Quarterly Sales Report",
            Items = new()
            {
                new Item { Index = 1, Name = "Product A" },
                new Item { Index = 2, Name = "Product B" },
                new Item { Index = 3, Name = "Product C" }
            }
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(report, model, "model");

        // Save the generated report.
        report.Save("Report.docx");
    }
}

// ---------------------------------------------------------------------
// Root data model referenced in the template as "model".
// ---------------------------------------------------------------------
public class ReportModel
{
    public string Title { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

// ---------------------------------------------------------------------
// Simple item class used in the foreach loop.
// ---------------------------------------------------------------------
public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
}
