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
}

public class PersonBookmark
{
    public string FullName { get; set; } = "";
    public string BookmarkName { get; set; } = "";
}

public class ReportModel
{
    public List<PersonBookmark> Bookmarks { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Sample data
        List<Person> persons = new()
        {
            new Person { FirstName = "John", LastName = "Doe" },
            new Person { FirstName = "Jane", LastName = "Smith" },
            new Person { FirstName = "Bob", LastName = "Johnson" }
        };

        // Use LINQ to create a collection with concatenated bookmark names
        List<PersonBookmark> bookmarks = persons
            .Select(p => new PersonBookmark
            {
                FullName = $"{p.FirstName} {p.LastName}",
                BookmarkName = $"{p.FirstName}{p.LastName}"
            })
            .ToList();

        ReportModel model = new() { Bookmarks = bookmarks };

        // Create template document programmatically
        string templatePath = "template.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        builder.Writeln("People List:");
        builder.Writeln("<<foreach [b in Bookmarks]>>");
        builder.Writeln("<<bookmark [b.BookmarkName]>>");
        builder.Writeln("<<[b.FullName]>>");
        builder.Writeln("<</bookmark>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load template and build report
        Document doc = new(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
