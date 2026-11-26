using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare output directory
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a sample image file (a small red dot PNG)
        string imagePath = Path.Combine(outputDir, "sample.png");
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAAFklEQVQoU2NkYGD4z0AEYBxVSFIAABcAAf+XK0cAAAAASUVORK5CYII=");
        File.WriteAllBytes(imagePath, pngBytes);

        // Build the DOCX template programmatically
        string templatePath = Path.Combine(outputDir, "template.docx");
        BuildTemplate(templatePath, imagePath);

        // Load the template
        Document templateDoc = new Document(templatePath);

        // Prepare sample data
        Order order = new()
        {
            CustomerName = "John Doe",
            Title = "Order Summary",
            HtmlSnippet = "<b>Thank you for your purchase!</b>",
            BookmarkName = "OrderSection",
            Url = "https://example.com/orders/12345",
            LinkText = "View Order Details",
            ColorName = "Green",
            Items = new()
            {
                new Item { Index = 1, Name = "Widget A", Price = 19.99m, ImagePath = imagePath },
                new Item { Index = 2, Name = "Widget B", Price = 29.99m, ImagePath = imagePath },
                new Item { Index = 3, Name = "Widget C", Price = 39.99m, ImagePath = imagePath }
            }
        };

        // Build the report
        ReportingEngine engine = new();
        engine.Options = ReportBuildOptions.None;
        bool success = engine.BuildReport(templateDoc, order, "order");

        // Save the result as PDF
        string pdfPath = Path.Combine(outputDir, "report.pdf");
        templateDoc.Save(pdfPath, SaveFormat.Pdf);
    }

    private static void BuildTemplate(string templatePath, string sampleImagePath)
    {
        Document doc = new();
        DocumentBuilder builder = new(doc);

        // Title
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Title;
        builder.Writeln("<<[order.Title]>>");

        // Customer name
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Customer: <<[order.CustomerName]>>");

        // HTML snippet
        builder.Writeln("<<[order.HtmlSnippet] -html>>");

        // Bookmark section
        builder.Writeln("<<bookmark [order.BookmarkName]>>");
        builder.Writeln("<<[order.Title]>>");
        builder.Writeln("<</bookmark>>");

        // Link
        builder.Writeln("<<link [order.Url] [order.LinkText]>>");

        // Text color example
        builder.Writeln("<<textColor [order.ColorName]>>Status: Active<</textColor>>");

        // Background color example
        builder.Writeln("<<backColor [\"LightGray\"]>>Background Text<</backColor>>");

        // Table with items (foreach inside the table)
        builder.Writeln("<<foreach [item in order.Items]>>");

        // Start table for each item (header + row)
        builder.StartTable();

        // Header row
        builder.InsertCell();
        builder.Writeln("Index");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Price");
        builder.InsertCell();
        builder.Writeln("Image");
        builder.EndRow();

        // Data row
        builder.InsertCell();
        builder.Writeln("<<[item.Index]>>");

        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");

        builder.InsertCell();
        builder.Writeln("<<[item.Price]>>");

        // Image inside a textbox
        builder.InsertCell();
        Shape textBox = builder.InsertShape(ShapeType.TextBox, 100, 100);
        builder.MoveTo(textBox.FirstParagraph);
        builder.Write("<<image [item.ImagePath] -fitSize>>");

        builder.EndRow();
        builder.EndTable();

        builder.Writeln("<</foreach>>");

        // Save the template
        doc.Save(templatePath);
    }

    // Data models
    public class Order
    {
        public string CustomerName { get; set; } = "";
        public string Title { get; set; } = "";
        public string HtmlSnippet { get; set; } = "";
        public string BookmarkName { get; set; } = "";
        public string Url { get; set; } = "";
        public string LinkText { get; set; } = "";
        public string ColorName { get; set; } = "";
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = "";
        public decimal Price { get; set; }
        public string ImagePath { get; set; } = "";
    }
}
