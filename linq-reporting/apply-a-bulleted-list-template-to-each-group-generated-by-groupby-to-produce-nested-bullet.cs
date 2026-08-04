using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data and group it by the first letter of the person's name.
        List<Person> people = new()
        {
            new() { Name = "Alice" },
            new() { Name = "Bob" },
            new() { Name = "Charlie" },
            new() { Name = "David" },
            new() { Name = "Eve" }
        };

        List<Group> groups = people
            .GroupBy(p => p.Name[0].ToString())
            .Select(g => new Group { Key = g.Key, Items = g.ToList() })
            .ToList();

        ReportModel model = new() { Groups = groups };

        // -----------------------------------------------------------------
        // Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document template = new();
        DocumentBuilder builder = new(template);

        // Create a bulleted list that will be used for both group headings and items.
        List bulletList = template.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = bulletList;

        // Outer foreach – iterate over groups.
        builder.Writeln("<<foreach [group in Model.Groups]>>");

        // Group heading – level 0 bullet.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("<<[group.Key]>>");

        // Inner foreach – iterate over persons inside the current group.
        builder.ListFormat.ListLevelNumber = 1; // level 1 bullet for items.
        builder.Writeln("<<foreach [person in group.Items]>>");
        builder.Writeln("<<[person.Name]>>");
        builder.Writeln("<</foreach>>");

        // End of outer foreach.
        builder.Writeln("<</foreach>>");

        // Optional: remove list formatting after the report is built.
        builder.ListFormat.List = null;

        // Save the template and reload it to satisfy the lifecycle rule.
        const string templatePath = "Template.docx";
        template.Save(templatePath);
        Document loadedTemplate = new(templatePath);

        // -----------------------------------------------------------------
        // Build the report using the ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new()
        {
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };
        engine.BuildReport(loadedTemplate, model, "Model");

        // Save the final document.
        const string outputPath = "Report.docx";
        loadedTemplate.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Group> Groups { get; set; } = new();
}

public class Group
{
    public string Key { get; set; } = "";
    public List<Person> Items { get; set; } = new();
}

public class Person
{
    public string Name { get; set; } = "";
}
