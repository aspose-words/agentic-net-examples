using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a blank document and a builder to insert LINQ Reporting tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Template: output a public member (Name) and a member we intend to restrict (Secret).
        builder.Writeln("Name: <<[model.Name]>>");
        builder.Writeln("Secret: <<[model.Secret]>>");

        // Prepare the data model.
        var model = new Person
        {
            Name = "John Doe",
            Secret = "TopSecret"
        };

        // Restrict access to the Person type (all its members become inaccessible in the template).
        // This must be done before any report is built.
        ReportingEngine.SetRestrictedTypes(typeof(Person));

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine
        {
            // Missing members (e.g., restricted members) will be treated as null instead of throwing.
            Options = ReportBuildOptions.AllowMissingMembers
        };

        // Build the report using the model as the root object named "model".
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save("Report.docx");
    }
}

// Simple data model with a public property (Name) and a property we intend to restrict (Secret).
public class Person
{
    public string Name { get; set; } = string.Empty;
    public string Secret { get; set; } = string.Empty;
}
