using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        ReportModel model = new()
        {
            Groups = new()
            {
                new Group
                {
                    Name = "Fruits",
                    Items = new() { "Apple", "Banana", "Cherry" }
                },
                new Group
                {
                    Name = "Vegetables",
                    Items = new() { "Carrot", "Lettuce", "Pepper" }
                }
            }
        };

        // Create the template document.
        Document template = new();
        DocumentBuilder builder = new(template);

        // Outer foreach over groups.
        builder.Writeln("<<foreach [group in model.Groups]>>");
        // Group name.
        builder.Writeln("<<[group.Name]>>");
        // Numbered list of items with restartNum before inner foreach.
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("<<restartNum>><<foreach [item in group.Items]>> <<[item]>> <</foreach>>");
        builder.ListFormat.RemoveNumbers();
        // End outer foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        const string templatePath = "template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        Document doc = new(templatePath);

        // Build the report.
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // Save the final document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}

// Data model classes.
public class ReportModel
{
    public List<Group> Groups { get; set; } = new();
}

public class Group
{
    public string Name { get; set; } = string.Empty;
    public List<string> Items { get; set; } = new();
}
