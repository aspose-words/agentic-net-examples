using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    static void Main()
    {
        // Create a template document with a placeholder that accesses a custom object's member.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Template syntax: <<[person.Name]>>
        builder.Writeln("Name: <<[person.Name]>>");

        // Define a custom data class instance.
        var person = new Person { Name = "John Doe", Age = 30 };

        // Restrict types that must not be accessible from the template (e.g., System.Type).
        ReportingEngine.SetRestrictedTypes(typeof(System.Type));

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine
        {
            // Treat missing members as null literals to avoid exceptions.
            Options = ReportBuildOptions.AllowMissingMembers
        };

        // Register external types that are safe to use in the template.
        // This enables static member access, type casts, etc.
        engine.KnownTypes.Add(typeof(DateTime));
        engine.KnownTypes.Add(typeof(Math));

        // Build the report using the template and the data source.
        var data = new DataSource { person = person };
        engine.BuildReport(doc, data);

        // Save the generated document.
        doc.Save("Report.docx");
    }

    // Simple POCO used as a data source for the template.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    // Wrapper class required because the reporting engine expects a visible type.
    public class DataSource
    {
        public Person person { get; set; }
    }
}
