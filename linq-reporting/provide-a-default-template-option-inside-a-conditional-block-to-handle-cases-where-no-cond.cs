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
        // Register code page provider for possible legacy encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 15 },
                new Person { Name = "Bob",   Age = 30 },
                new Person { Name = "Carol", Age = 70 },
                new Person { Name = "Dave",  Age = -1 } // Unexpected age to trigger default.
            }
        };

        // Create a template document with LINQ Reporting tags.
        const string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("<<foreach [person in Persons]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Category: ");
        builder.Writeln("<<if [person.Age < 18]>>Minor<</if>>");
        builder.Writeln("<<if [person.Age >= 65]>>Senior<</if>>");
        builder.Writeln("<<if [person.Age >= 18 && person.Age < 65]>>Adult<</if>>");
        // Default option when none of the above conditions are true.
        builder.Writeln("<<if [person.Age < 0]>>Unknown<</if>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template and build the report.
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}

// Data model classes.
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}
