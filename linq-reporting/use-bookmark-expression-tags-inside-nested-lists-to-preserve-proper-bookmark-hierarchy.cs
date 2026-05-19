using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a blank document that will serve as the template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin the outer foreach loop over categories.
        builder.Writeln("<<foreach [category in Categories]>>");

        // Create a numbered list for categories.
        builder.ListFormat.List = doc.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.ListLevelNumber = 0;

        // Each category name is wrapped in a bookmark tag.
        builder.Writeln("<<bookmark [category.Name]>><<[category.Name]>><</bookmark>>");

        // Begin the inner foreach loop over products of the current category.
        builder.Writeln("<<foreach [product in category.Products]>>");

        // Indent the inner list (level 1) and use a bullet for each product.
        builder.ListFormat.ListLevelNumber = 1;
        builder.Writeln("- <<bookmark [product.Name]>><<[product.Name]>><</bookmark>>");

        // End inner foreach.
        builder.Writeln("<</foreach>>");

        // Reset list level back to the outer level.
        builder.ListFormat.ListLevelNumber = 0;

        // End outer foreach.
        builder.Writeln("<</foreach>>");

        // Remove list formatting from subsequent paragraphs.
        builder.ListFormat.List = null;

        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            Categories = new List<Category>
            {
                new Category
                {
                    Name = "Fruits",
                    Products = new List<Product>
                    {
                        new Product { Name = "Apple" },
                        new Product { Name = "Banana" }
                    }
                },
                new Category
                {
                    Name = "Vegetables",
                    Products = new List<Product>
                    {
                        new Product { Name = "Carrot" },
                        new Product { Name = "Tomato" }
                    }
                }
            }
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}

// Root data model.
public class ReportModel
{
    public List<Category> Categories { get; set; } = new();
}

// Category containing a list of products.
public class Category
{
    public string Name { get; set; } = string.Empty;
    public List<Product> Products { get; set; } = new();
}

// Simple product model.
public class Product
{
    public string Name { get; set; } = string.Empty;
}
