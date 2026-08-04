using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create a template document that contains a reference to a missing field.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Name: <<[person.Name]>>");
        // The Age property does NOT exist in the Person class.
        builder.Writeln("Age: <<[person.Age]>>");
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template for reporting.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the data source. Only Name is defined.
        // -----------------------------------------------------------------
        var person = new Person { Name = "John Doe" };

        // -----------------------------------------------------------------
        // 4. Configure the ReportingEngine to treat missing members as null.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.AllowMissingMembers;
        engine.MissingMemberMessage = "N/A";

        // -----------------------------------------------------------------
        // 5. Build the report. The root object name in the template is "person".
        // -----------------------------------------------------------------
        engine.BuildReport(reportDoc, person, "person");

        // -----------------------------------------------------------------
        // 6. Save the generated report.
        // -----------------------------------------------------------------
        reportDoc.Save(outputPath);
    }

    // Simple data model with only a Name property.
    public class Person
    {
        public string Name { get; set; } = "";
        // No Age property – it will be treated as null by the engine.
    }
}
