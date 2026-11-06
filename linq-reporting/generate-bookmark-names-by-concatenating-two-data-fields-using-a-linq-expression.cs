using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string FirstName { get; set; } = "";
    public string LastName { get; set; } = "";
    public string BookmarkName { get; set; } = "";
}

public class Model
{
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Paths for template and output documents.
        string templatePath = "Template.docx";
        string outputPath = "Report.docx";

        // Create the template document with LINQ Reporting tags.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin a foreach loop over the collection "Persons".
        builder.Writeln("<<foreach [person in Persons]>>");

        // Define a bookmark whose name is the concatenated fields.
        builder.Writeln("<<bookmark [person.BookmarkName]>>");
        // Bookmark content: display the person's full name.
        builder.Writeln("<<[person.FirstName]>> <<[person.LastName]>>");
        builder.Writeln("<</bookmark>>");

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Sample data source.
        var rawPersons = new List<Person>
        {
            new Person { FirstName = "John", LastName = "Doe" },
            new Person { FirstName = "Jane", LastName = "Smith" },
            new Person { FirstName = "Alice", LastName = "Johnson" }
        };

        // Use LINQ to create bookmark names by concatenating first and last names.
        var personsWithBookmarks = rawPersons
            .Select(p => new Person
            {
                FirstName = p.FirstName,
                LastName = p.LastName,
                BookmarkName = $"{p.FirstName}{p.LastName}"
            })
            .ToList();

        // Wrap the collection in a root model object.
        var model = new Model { Persons = personsWithBookmarks };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        reportDoc.Save(outputPath);
    }
}
