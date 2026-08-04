using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Header section ----------
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("<<[header.Title]>>");
        builder.Writeln("Date: <<[header.Date]>>");

        // ---------- Body section ----------
        builder.MoveToDocumentEnd();
        builder.Writeln("<<foreach [item in body.Items]>>");
        builder.Writeln("Product: <<[item.Name]>> - Qty: <<[item.Quantity]>>");
        builder.Writeln("<</foreach>>");

        // ---------- Footer section ----------
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("Page <<[footer.PageNumber]>> of <<[footer.TotalPages]>>");

        // Prepare data sources.
        HeaderModel header = new HeaderModel
        {
            Title = "Sales Report",
            Date = DateTime.Now.ToString("d")
        };

        BodyModel body = new BodyModel
        {
            Items = new()
            {
                new Item { Name = "Apple", Quantity = 10 },
                new Item { Name = "Banana", Quantity = 20 },
                new Item { Name = "Cherry", Quantity = 15 }
            }
        };

        FooterModel footer = new FooterModel
        {
            PageNumber = 1,
            TotalPages = 1
        };

        // Build the report using three separate data sources.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;
        engine.BuildReport(doc,
            new object[] { header, body, footer },
            new string[] { "header", "body", "footer" });

        // Save the generated report.
        doc.Save("MultiSectionReport.docx");
    }
}

// Header data model.
public class HeaderModel
{
    public string Title { get; set; } = string.Empty;
    public string Date { get; set; } = string.Empty;
}

// Body data model containing a collection of items.
public class BodyModel
{
    public List<Item> Items { get; set; } = new();
}

// Individual item used in the body collection.
public class Item
{
    public string Name { get; set; } = string.Empty;
    public int Quantity { get; set; }
}

// Footer data model.
public class FooterModel
{
    public int PageNumber { get; set; }
    public int TotalPages { get; set; }
}
