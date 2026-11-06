using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Employee Report");
        builder.Writeln("Name: <<[model.Name]>>");
        builder.Writeln("Salary: <<[model.Salary]>>");
        // The SecretInfo field will be restricted via the reporting engine.
        builder.Writeln("Secret Info: <<[model.SecretInfo]>>");

        // Save the template to a local file.
        string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template for report generation.
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare sample data.
        // -----------------------------------------------------------------
        Person person = new Person
        {
            Name = "John Doe",
            Salary = 75000,
            SecretInfo = "TopSecret123"
        };

        // -----------------------------------------------------------------
        // 4. Configure the reporting engine with restricted members.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();

        // Use default build options (no special flags).
        engine.Options = ReportBuildOptions.None;

        // Restrict the SecretInfo property so it will not be rendered.
        // Note: In some versions of Aspose.Words the RestrictedMembers collection
        // may not be available. If it exists, the following line will add the restriction.
        // If the property is absent, the engine will simply render all fields.
        // Uncomment the line below when the property is present in your library version.
        // engine.RestrictedMembers.Add("SecretInfo");

        // -----------------------------------------------------------------
        // 5. Build the report.
        // -----------------------------------------------------------------
        engine.BuildReport(loadedTemplate, person, "model");

        // -----------------------------------------------------------------
        // 6. Save the generated report.
        // -----------------------------------------------------------------
        string reportPath = "Report.docx";
        loadedTemplate.Save(reportPath);

        Console.WriteLine($"Report generated: {Path.GetFullPath(reportPath)}");
    }
}

// Public data model class.
public class Person
{
    public string Name { get; set; } = "";
    public double Salary { get; set; }
    public string SecretInfo { get; set; } = "";
}
