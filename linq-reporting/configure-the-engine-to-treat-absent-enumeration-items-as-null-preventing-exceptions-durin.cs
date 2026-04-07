using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class IncompletePerson
{
    public string Name { get; set; } = "";
    // No Age property – this will be missing in the template.
}

public class ReportModel
{
    public List<object> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // 1. Create a template document programmatically.
        var templatePath = "Template.docx";
        var builder = new DocumentBuilder();
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");
        builder.Document.Save(templatePath);

        // 2. Load the template back from disk.
        var doc = new Document(templatePath);

        // 3. Prepare sample data with a missing member (Age) in one item.
        var model = new ReportModel();
        model.Persons.Add(new Person { Name = "Alice", Age = 30 });
        model.Persons.Add(new IncompletePerson { Name = "Bob" }); // Age missing

        // 4. Configure the ReportingEngine to treat missing members as null.
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers
        };
        // Optional: customize the message shown for missing members.
        engine.MissingMemberMessage = "";

        // 5. Build the report.
        engine.BuildReport(doc, model, "model");

        // 6. Save the generated report.
        doc.Save("Report.docx");
    }
}
