using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required by Aspose.Words for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data model
        ReportModel model = new()
        {
            Categories = new()
            {
                new()
                {
                    Name = "Fruits",
                    Bookmark = "Bookmark_Fruits",
                    Products = new()
                    {
                        new() { Name = "Apple", Bookmark = "Bookmark_Apple" },
                        new() { Name = "Banana", Bookmark = "Bookmark_Banana" }
                    }
                },
                new()
                {
                    Name = "Vegetables",
                    Bookmark = "Bookmark_Vegetables",
                    Products = new()
                    {
                        new() { Name = "Carrot", Bookmark = "Bookmark_Carrot" },
                        new() { Name = "Lettuce", Bookmark = "Bookmark_Lettuce" }
                    }
                }
            }
        };

        // Create template document programmatically
        string templatePath = "template.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        builder.Writeln("Report with nested bookmarks:");
        builder.Writeln();

        // Outer foreach for categories
        builder.Writeln("<<foreach [cat in Model.Categories]>>");
        // Category bookmark
        builder.Writeln("<<bookmark [cat.Bookmark]>>");
        builder.Writeln("<<[cat.Name]>>");
        builder.Writeln("<</bookmark>>");

        // Inner foreach for products
        builder.Writeln("<<foreach [prod in cat.Products]>>");
        builder.Writeln("    - <<bookmark [prod.Bookmark]>><<[prod.Name]>> <</bookmark>>");
        builder.Writeln("<</foreach>>");

        // End outer foreach
        builder.Writeln("<</foreach>>");

        // Save the template
        templateDoc.Save(templatePath);

        // Load the template for reporting
        Document reportDoc = new(templatePath);

        // Build the report using LINQ Reporting Engine
        ReportingEngine engine = new();
        engine.BuildReport(reportDoc, model, "Model");

        // Save the generated report
        string outputPath = "output.docx";
        reportDoc.Save(outputPath);
    }
}

// Data model classes
public class ReportModel
{
    public List<Category> Categories { get; set; } = new();
}

public class Category
{
    public string Name { get; set; } = "";
    public string Bookmark { get; set; } = "";
    public List<Product> Products { get; set; } = new();
}

public class Product
{
    public string Name { get; set; } = "";
    public string Bookmark { get; set; } = "";
}
