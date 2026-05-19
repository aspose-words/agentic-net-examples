using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    // Initialize properties to avoid nullable warnings.
    public string Name { get; set; } = string.Empty;
    public string Secret { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        string templatePath = "Template.docx";
        string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // The <<restrictMembers>> tag is not required for the current Aspose.Words version.
        // Member access is restricted via ReportingEngine.SetRestrictedTypes below.
        // builder.Writeln("<<restrictMembers>>"); // Removed to avoid parsing error.

        // Allowed member access.
        builder.Writeln("Name: <<[person.Name]>>");

        // This member will be blocked because the Person type is restricted.
        builder.Writeln("Secret: <<[person.Secret]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template for reporting.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Restrict access to the Person type (all its members become inaccessible).
        // -----------------------------------------------------------------
        ReportingEngine.SetRestrictedTypes(typeof(Person));

        // -----------------------------------------------------------------
        // 4. Configure the reporting engine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine
        {
            // Allow missing members so that restricted members are treated as null.
            Options = ReportBuildOptions.AllowMissingMembers
        };

        // Optional: customize the message shown for missing members.
        engine.MissingMemberMessage = string.Empty;

        // -----------------------------------------------------------------
        // 5. Prepare the data source.
        // -----------------------------------------------------------------
        Person person = new Person { Name = "John Doe", Secret = "TopSecret" };

        // -----------------------------------------------------------------
        // 6. Build the report. The root object name must match the tag reference.
        // -----------------------------------------------------------------
        engine.BuildReport(doc, person, "person");

        // -----------------------------------------------------------------
        // 7. Save the generated report.
        // -----------------------------------------------------------------
        doc.Save(reportPath);

        // Indicate completion.
        Console.WriteLine($"Report generated: {Path.GetFullPath(reportPath)}");
    }
}
