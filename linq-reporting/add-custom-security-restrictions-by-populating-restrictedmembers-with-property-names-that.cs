using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    // Public properties referenced by the template.
    public string Name { get; set; } = string.Empty;
    public string Secret { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";

        // -------------------------------------------------
        // Create the template document programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert LINQ Reporting tags that reference the model's members.
        builder.Writeln("Name: <<[model.Name]>>");
        builder.Writeln("Secret: <<[model.Secret]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // Load the template for report generation.
        // -------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // Sample data model.
        Person model = new Person
        {
            Name = "John Doe",
            Secret = "TopSecretInformation"
        };

        // -------------------------------------------------
        // Configure the ReportingEngine with custom security restrictions.
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine();

        // Restrict access to the Person type (all its members). This is the
        // available API for setting restricted types. To avoid runtime errors
        // when a restricted member is referenced in the template, enable the
        // AllowMissingMembers option so missing members are treated as null.
        ReportingEngine.SetRestrictedTypes(typeof(Person));
        engine.Options = ReportBuildOptions.AllowMissingMembers;

        // Build the report using the loaded template and the data model.
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the final report.
        loadedTemplate.Save(outputPath);
    }
}
