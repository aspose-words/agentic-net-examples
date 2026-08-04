using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a template document with LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("<<foreach [item in Persons]>>");
        builder.Writeln("Name: <<[item.Name]>>");
        builder.Writeln("Age: <<[item.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save and reload the template to simulate a real file workflow.
        const string templatePath = "Template.docx";
        template.Save(templatePath);
        Document doc = new Document(templatePath);

        // Prepare the data model. The collection contains objects with and without the 'Age' member.
        ReportModel model = new ReportModel
        {
            Persons = new List<object>
            {
                new Person { Name = "John", Age = 30 },
                new Dummy { Name = "Jane" },          // Missing 'Age' member.
                new Person { Name = "Bob", Age = 25 }
            }
        };

        // Configure the reporting engine to treat missing members as null.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers,
            MissingMemberMessage = "" // Optional: suppress custom missing member text.
        };

        // Build the report using the model as the root object named "model".
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);

        Console.WriteLine($"Report generated: {outputPath}");
    }
}

// Root data model.
public class ReportModel
{
    public List<object> Persons { get; set; } = new();
}

// Object with both Name and Age members.
public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

// Object missing the Age member.
public class Dummy
{
    public string Name { get; set; } = "";
}
