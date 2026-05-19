using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}

public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Prepare sample data.
        // -----------------------------------------------------------------
        var model = new ReportModel();
        model.Persons.Add(new Person { Name = "John Doe", Age = 30 });
        model.Persons.Add(new Person { Name = "Jane Smith", Age = 25 });

        // -----------------------------------------------------------------
        // 2. Build the template document programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("People Report");
        builder.Writeln(); // empty line

        // Start a foreach block that iterates over the Persons collection.
        builder.Writeln("<<foreach [person in Persons]>>");

        // Create a table that will be repeated for each person.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Age");
        builder.EndRow();

        // Data row – the engine will replace the tags with actual values.
        builder.InsertCell();
        builder.Writeln("<<[person.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[person.Age]>>");
        builder.EndRow();

        builder.EndTable();

        // Close the foreach block.
        builder.Writeln("<</foreach>>");

        // -----------------------------------------------------------------
        // 3. Build the report using the LINQ Reporting engine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // Use a data source name ("model") so that the template can reference it if needed.
        engine.BuildReport(template, model, "model");

        // -----------------------------------------------------------------
        // 4. Save the generated document.
        // -----------------------------------------------------------------
        template.Save("Report.docx");
    }
}
