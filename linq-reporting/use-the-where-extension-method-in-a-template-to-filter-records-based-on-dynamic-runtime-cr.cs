using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
    public int MinAge { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Sample data.
        var model = new ReportModel
        {
            MinAge = 30,
            Persons = new()
            {
                new Person { Name = "Alice", Age = 25 },
                new Person { Name = "Bob", Age = 35 },
                new Person { Name = "Charlie", Age = 40 }
            }
        };

        // -----------------------------------------------------------------
        // Create the template document programmatically.
        // -----------------------------------------------------------------
        const string templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Header line – reference the root object (model) for MinAge.
        builder.Writeln("Persons with Age >= <<[model.MinAge]>>:");

        // Foreach loop over the collection.
        builder.Writeln("<<foreach [p in Persons]>>");

        // Conditional output – compare the person's age with the root MinAge.
        builder.Writeln("<<if [p.Age >= model.MinAge]>>Name: <<[p.Name]>>, Age: <<[p.Age]>> <</if>>");

        // End of foreach.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        doc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and generate the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the model as the root data source named "model".
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
