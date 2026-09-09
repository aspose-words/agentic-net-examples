using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Root data model for the report.
    public class ReportModel
    {
        public List<Group> Groups { get; set; } = new();
    }

    // Represents a group produced by GroupBy.
    public class Group
    {
        public string Key { get; set; } = string.Empty;          // Group key (e.g., Department)
        public List<Person> Items { get; set; } = new();         // Items belonging to the group
    }

    // Sample data entity.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public string Department { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create sample data and group it.
            // -----------------------------------------------------------------
            List<Person> persons = new()
            {
                new Person { Name = "Alice",   Department = "HR" },
                new Person { Name = "Bob",     Department = "IT" },
                new Person { Name = "Charlie", Department = "HR" },
                new Person { Name = "David",   Department = "Finance" },
                new Person { Name = "Eve",     Department = "IT" }
            };

            // Group by Department and project to the model structure.
            ReportModel model = new()
            {
                Groups = persons
                    .GroupBy(p => p.Department)
                    .Select(g => new Group
                    {
                        Key = g.Key,
                        Items = g.ToList()
                    })
                    .ToList()
            };

            // -----------------------------------------------------------------
            // 2. Build the LINQ Reporting template programmatically.
            // -----------------------------------------------------------------
            string templatePath = "Template.docx";

            // Create a blank document and a builder.
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Create a bulleted list template.
            List bulletList = templateDoc.Lists.Add(ListTemplate.BulletDefault);
            builder.ListFormat.List = bulletList;

            // First level – group name.
            builder.ListFormat.ListLevelNumber = 0;
            builder.Writeln("<<foreach [g in Groups]>>");
            builder.Writeln("<<[g.Key]>>");

            // Second level – items inside the group.
            builder.ListFormat.ListLevelNumber = 1;
            builder.Writeln("<<foreach [p in g.Items]>>");
            builder.Writeln("<<[p.Name]>>");
            builder.Writeln("<</foreach>>");

            // End of outer foreach.
            builder.ListFormat.ListLevelNumber = 0;
            builder.Writeln("<</foreach>>");

            // Clean up list formatting.
            builder.ListFormat.RemoveNumbers();

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 3. Load the template and generate the report.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            // Build the report using the root object name "model".
            engine.BuildReport(reportDoc, model, "model");

            // Save the final document.
            string outputPath = "Report.docx";
            reportDoc.Save(outputPath);

            Console.WriteLine($"Report generated successfully: {Path.GetFullPath(outputPath)}");
        }
    }
}
