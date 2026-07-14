using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class People
{
    public List<Person> Person { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for UTF‑8 support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        const string templatePath = "template.docx";
        const string outputPath = "report.docx";

        // -----------------------------------------------------------------
        // 1. Build a Word template that contains LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("People Report");
        builder.Writeln("----------------");

        // The data source will be referenced by the name "people".
        // The collection of Person elements is accessed as people.Person.
        builder.Writeln("<<foreach [p in people.Person]>>");
        builder.Writeln("- Name: <<[p.Name]>>");
        builder.Writeln("  Age : <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Create sample data model.
        // -----------------------------------------------------------------
        var people = new People();
        people.Person.Add(new Person { Name = "John Doe", Age = 30 });
        people.Person.Add(new Person { Name = "Jane Smith", Age = 25 });
        people.Person.Add(new Person { Name = "Bob Johnson", Age = 40 });

        // -----------------------------------------------------------------
        // 3. Load the template and bind the object data source.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // Build the report. The third argument is the name used inside the template.
        engine.BuildReport(reportDoc, people, "people");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        reportDoc.Save(outputPath);
    }
}
