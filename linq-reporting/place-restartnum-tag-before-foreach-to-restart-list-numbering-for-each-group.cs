using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Groups = new List<Group>
            {
                new Group
                {
                    Name = "Group A",
                    Items = new List<string> { "Item A1", "Item A2", "Item A3" }
                },
                new Group
                {
                    Name = "Group B",
                    Items = new List<string> { "Item B1", "Item B2" }
                }
            }
        };

        // Create the template document.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Begin outer foreach over groups.
        builder.Writeln("<<foreach [group in Model.Groups]>>");
        // Write group name.
        builder.Writeln("<<[group.Name]>>");

        // Start a numbered list for items within the group.
        builder.ListFormat.ApplyNumberDefault();
        // Restart numbering for each group and iterate over items.
        builder.Writeln("<<restartNum>><<foreach [item in group.Items]>> <<[item]>> <</foreach>>");
        // End the numbered list.
        builder.ListFormat.RemoveNumbers();

        // End outer foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report.
        engine.BuildReport(reportDoc, model, "Model");

        // Save the final report.
        reportDoc.Save("Report.docx");
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
