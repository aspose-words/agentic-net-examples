using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        var templatePath = "Template.docx";
        var builder = new DocumentBuilder();

        // Loop over the collection of persons.
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("Name: <<[p.Name]>>");

        // Show status for adults.
        builder.Writeln("<<if [p.IsAdult()]>>Status: Adult<</if>>");

        // Show status for minors – use a comparison instead of the '!' operator,
        // because the engine cannot apply '!' to an object.
        builder.Writeln("<<if [p.IsAdult() == false]>>Status: Minor<</if>>");

        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        builder.Document.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template for reporting.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the data model.
        // -----------------------------------------------------------------
        var model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = 15 }
            }
        };

        // -----------------------------------------------------------------
        // 4. Configure the reporting engine.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine
        {
            // Allow missing members so that the engine treats absent members as null.
            Options = ReportBuildOptions.AllowMissingMembers
        };

        // Register the class that contains the extension method.
        engine.KnownTypes.Add(typeof(MyExtensions));

        // -----------------------------------------------------------------
        // 5. Build the report.
        // -----------------------------------------------------------------
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 6. Save the generated report.
        // -----------------------------------------------------------------
        doc.Save("Report.docx");
    }
}

// ---------------------------------------------------------------------
// Root data model.
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

// ---------------------------------------------------------------------
// Simple data entity.
// ---------------------------------------------------------------------
public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}

// ---------------------------------------------------------------------
// Extension methods usable in templates.
// ---------------------------------------------------------------------
public static class MyExtensions
{
    // Returns true if the person is 18 or older.
    public static bool IsAdult(this Person person) => person.Age >= 18;
}
