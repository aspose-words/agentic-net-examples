using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

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
        // Prepare sample data – some entries will produce empty paragraphs.
        var model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "", Age = 0 },          // Will result in an empty paragraph.
                new Person { Name = "Bob", Age = 25 },
                new Person { Name = null!, Age = 0 }        // Will also result in an empty paragraph.
            }
        };

        // Build the template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // LINQ Reporting foreach loop with a conditional that may output nothing.
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("<<if [p.Name != null && p.Name != \"\"]>><<[p.Name]>> - <<[p.Age]>> <</if>>");
        builder.Writeln("<</foreach>>");

        // Configure the reporting engine to remove empty paragraphs after processing.
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };

        // Build the report using the model as the root data source named "model".
        engine.BuildReport(doc, model, "model");

        // Save the final document.
        doc.Save("Report.docx");
    }
}
