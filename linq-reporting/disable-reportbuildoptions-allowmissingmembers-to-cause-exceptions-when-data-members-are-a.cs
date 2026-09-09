using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Simple data model with only Name property.
    public class Person
    {
        public string Name { get; set; } = "John Doe";
    }

    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a LINQ Reporting tag that references a missing member (Age).
        // The template expects a property called Age, which does not exist in Person.
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>"); // <-- missing member

        // Save the template (optional, just to see the generated document).
        const string templatePath = "ReportTemplate.docx";
        doc.Save(templatePath);

        // Prepare the data source.
        var person = new Person();

        // Create the reporting engine without AllowMissingMembers flag.
        ReportingEngine engine = new ReportingEngine();
        // Ensure no special options are set (default is ReportBuildOptions.None).
        engine.Options = ReportBuildOptions.None;

        try
        {
            // Build the report. This should throw an exception because Age is missing.
            engine.BuildReport(doc, person, "person");
            // If no exception, save the resulting document.
            doc.Save("ReportResult.docx");
            Console.WriteLine("Report built successfully (unexpected).");
        }
        catch (Exception ex)
        {
            // Expected path: an exception is thrown for the missing member.
            Console.WriteLine("Exception caught as expected:");
            Console.WriteLine(ex.Message);
        }
    }
}
