using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // 1. Create the template document with LINQ Reporting tags.
        var templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Start a foreach loop over the collection "Persons".
        builder.Writeln("<<foreach [p in Persons]>>");

        // Create a bookmark whose name is the concatenation of FirstName and LastName.
        // The expression inside the bookmark tag evaluates for each item.
        builder.Writeln("<<bookmark [p.FirstName + \"_\" + p.LastName]>>");
        // Content inside the bookmark (optional, just for demonstration).
        builder.Writeln("<<[p.FirstName]>> <<[p.LastName]>>");
        builder.Writeln("<</bookmark>>");

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // 2. Load the template document for report generation.
        var reportDoc = new Document(templatePath);

        // 3. Prepare the data source.
        var model = new ReportModel
        {
            Persons = new()
            {
                new Person { FirstName = "John", LastName = "Doe" },
                new Person { FirstName = "Jane", LastName = "Smith" },
                new Person { FirstName = "Alice", LastName = "Johnson" }
            }
        };

        // 4. Build the report using the ReportingEngine.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // 5. Save the generated report.
        reportDoc.Save("Report.docx");
    }
}

// Root data model referenced in the template as "model".
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

// Simple data entity with two fields to be concatenated.
public class Person
{
    public string FirstName { get; set; } = string.Empty;
    public string LastName { get; set; } = string.Empty;
}
