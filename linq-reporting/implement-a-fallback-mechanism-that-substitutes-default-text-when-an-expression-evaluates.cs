using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    // Nullable property to allow null values.
    public string? Name { get; set; }
}

public class ReportModel
{
    // Collection of persons to iterate over.
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Create a template document programmatically.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin a foreach loop over the Persons collection.
        builder.Writeln("<<foreach [p in Persons]>>");

        // If the Name is not null, display it; otherwise, show default text.
        builder.Writeln(
            "<<if [p.Name != null]>>Name: <<[p.Name]>> <</if>>" +
            "<<if [p.Name == null]>>Name: (no name)<</if>>");

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Prepare sample data with one person having a name and another with null.
        ReportModel model = new ReportModel();
        model.Persons.Add(new Person { Name = "Alice" });
        model.Persons.Add(new Person { Name = null });

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        // No special options are required for this example.
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
