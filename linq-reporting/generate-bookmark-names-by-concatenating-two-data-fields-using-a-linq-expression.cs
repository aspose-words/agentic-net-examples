using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings in Aspose.Words).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Create a template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Begin a foreach loop over the collection "Persons".
        builder.Writeln("<<foreach [person in Persons]>>");

        // Create a bookmark whose name is the concatenation of FirstName and LastName.
        // The expression inside the bookmark tag is evaluated for each item.
        builder.Writeln("<<bookmark [person.FirstName + \" \" + person.LastName]>>");

        // The visible content of the bookmark (optional, just for demonstration).
        builder.Writeln("<<[person.FirstName]>> <<[person.LastName]>>");

        // Close the bookmark and the foreach block.
        builder.Writeln("<</bookmark>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk (required before building the report).
        string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        Document report = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { FirstName = "John",  LastName = "Doe" },
                new Person { FirstName = "Jane",  LastName = "Smith" },
                new Person { FirstName = "Alice", LastName = "Johnson" }
            }
        };

        // Create the reporting engine and build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(report, model, "model");

        // -----------------------------------------------------------------
        // 3. Save the generated report.
        // -----------------------------------------------------------------
        string outputPath = "Report.docx";
        report.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes (public, non‑nullable properties are initialized).
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

public class Person
{
    public string FirstName { get; set; } = string.Empty;
    public string LastName  { get; set; } = string.Empty;
}
