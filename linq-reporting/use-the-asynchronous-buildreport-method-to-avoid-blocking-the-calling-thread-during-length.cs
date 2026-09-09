using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    // Sample property used in the template.
    public string Name { get; set; } = "John Doe";
}

public class Model
{
    // Collection that will be iterated in the template.
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    // Async entry point to avoid blocking the calling thread.
    public static async Task Main()
    {
        // -------------------------------------------------
        // 1. Create a template document programmatically.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple LINQ Reporting tag that iterates over the collection.
        builder.Writeln("<<foreach [person in Persons]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("<</foreach>>");

        // -------------------------------------------------
        // 2. Prepare sample data.
        // -------------------------------------------------
        Model model = new Model();
        model.Persons.Add(new Person { Name = "Alice" });
        model.Persons.Add(new Person { Name = "Bob" });
        model.Persons.Add(new Person { Name = "Charlie" });

        // -------------------------------------------------
        // 3. Build the report asynchronously.
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine();

        // Run the synchronous BuildReport inside Task.Run to keep the UI thread responsive.
        bool success = await Task.Run(() => engine.BuildReport(doc, model, "model"));

        // -------------------------------------------------
        // 4. Save the generated report.
        // -------------------------------------------------
        if (success)
        {
            doc.Save("Report_Output.docx");
            Console.WriteLine("Report generated successfully.");
        }
        else
        {
            Console.WriteLine("Report generation failed.");
        }
    }
}
