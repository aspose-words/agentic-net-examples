using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;

public class Person
{
    public string Department { get; set; } = "";
    public string Name { get; set; } = "";
}

public class Group
{
    public string Key { get; set; } = "";
    public List<string> Items { get; set; } = new();
}

public class ReportModel
{
    public List<Group> Groups { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var persons = new List<Person>
        {
            new Person { Department = "Engineering", Name = "Alice" },
            new Person { Department = "Engineering", Name = "Bob" },
            new Person { Department = "HR", Name = "Carol" },
            new Person { Department = "HR", Name = "Dave" },
            new Person { Department = "Marketing", Name = "Eve" }
        };

        // Group persons by department and map to the model used by the reporting engine.
        var model = new ReportModel
        {
            Groups = persons
                .GroupBy(p => p.Department)
                .Select(g => new Group
                {
                    Key = g.Key,
                    Items = g.Select(p => p.Name).ToList()
                })
                .ToList()
        };

        // -----------------------------------------------------------------
        // Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Create a bulleted list style that will be applied to the paragraphs.
        List bulletList = templateDoc.Lists.Add(ListTemplate.BulletDefault);

        // Begin outer foreach over groups.
        builder.Writeln("<<foreach [g in model.Groups]>>");

        // Apply the list to the group title (first level bullet).
        builder.ListFormat.List = bulletList;
        builder.Writeln("<<[g.Key]>>");

        // Indent to second level for items inside the group.
        builder.ListFormat.ListIndent();

        // Begin inner foreach over items of the current group.
        builder.Writeln("<<foreach [p in g.Items]>>");
        builder.Writeln("<<[p]>>");
        builder.Writeln("<</foreach>>");

        // Outdent back to first level and close the outer foreach.
        builder.ListFormat.ListOutdent();
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // No special options required.

        // Build the report using the model; the root object name is "model".
        bool success = engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string reportPath = "Report.docx";
        reportDoc.Save(reportPath);

        // Optional: indicate success.
        Console.WriteLine(success ? "Report generated successfully." : "Report generation failed.");
    }
}
