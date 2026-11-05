using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required on .NET Core/5+)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data
        var model = new ReportModel
        {
            Categories = new List<Category>
            {
                new Category
                {
                    Name = "Fruits",
                    Items = new List<string> { "Apple", "Banana", "Cherry" }
                },
                new Category
                {
                    Name = "Vegetables",
                    Items = new List<string> { "Carrot", "Lettuce", "Tomato" }
                }
            }
        };

        // Serialize to a local JSON file (optional, demonstrates JSON source)
        const string jsonPath = "data.json";
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(model, Formatting.Indented));

        // Create the LINQ Reporting template programmatically
        const string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Title
        builder.Writeln("Category Report");
        builder.Writeln();

        // Outer foreach for categories
        builder.Writeln("<<foreach [cat in Categories]>>");
        // Category name as a bullet (using a dash)
        builder.Writeln("- <<[cat.Name]>>");
        // Inner foreach for items within a category
        builder.Writeln("<<foreach [itm in cat.Items]>>");
        // Item bullet, indented with a tab for visual nesting
        builder.Writeln("\t- <<[itm]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        // Save the template
        templateDoc.Save(templatePath);

        // Load the template for report generation
        var doc = new Document(templatePath);

        // Build the report using the model as the root object named "model"
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the final document
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}

// Root model class
public class ReportModel
{
    public List<Category> Categories { get; set; } = new();
}

// Category class with a collection of items
public class Category
{
    public string Name { get; set; } = "";
    public List<string> Items { get; set; } = new();
}
