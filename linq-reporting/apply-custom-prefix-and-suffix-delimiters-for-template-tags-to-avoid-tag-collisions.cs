using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class Model
{
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings).
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // ---------- Create the template document ----------
        const string templateFile = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Use the default LINQ Reporting delimiters << and >>.
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("<<[p.Name]>> - <<[p.Age]>> years old");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templateFile);

        // ---------- Load the template ----------
        var doc = new Document(templateFile);

        // ---------- Prepare sample data ----------
        var model = new Model
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = 25 }
            }
        };

        // ---------- Configure ReportingEngine ----------
        var engine = new ReportingEngine();

        // ---------- Build the report ----------
        engine.BuildReport(doc, model, "model");

        // ---------- Save the generated report ----------
        doc.Save("Report.docx");
    }
}
