using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    // Name may be null to demonstrate fallback.
    public string? Name { get; set; }
}

public class ReportModel
{
    // Wrapper object referenced in the template as "model".
    public Person Person { get; set; } = new Person();
}

public class Program
{
    public static void Main()
    {
        // Ensure the working directory exists.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(workDir);

        // Paths for the template and the generated report.
        string templatePath = Path.Combine(workDir, "template.docx");
        string reportPath = Path.Combine(workDir, "report.docx");

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Write a label.
        builder.Writeln("Customer Name:");

        // If the Name property is not null, output its value.
        builder.Writeln("<<if [model.Person.Name != null]>>");
        builder.Writeln("<<[model.Person.Name]>>");
        builder.Writeln("<</if>>");

        // If the Name property is null, output a default placeholder.
        builder.Writeln("<<if [model.Person.Name == null]>>");
        builder.Writeln("[No Name Provided]");
        builder.Writeln("<</if>>");

        // Save the template to disk (required before BuildReport).
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // Prepare the data model with a null Name to trigger the fallback.
        ReportModel model = new ReportModel
        {
            Person = new Person
            {
                Name = null // Intentionally null.
            }
        };

        // Create the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Build the report using the model; the root object name is "model".
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(reportPath);

        // -----------------------------------------------------------------
        // 3. Demonstrate the same template with a non‑null value.
        // -----------------------------------------------------------------
        Document docWithName = new Document(templatePath);
        ReportModel modelWithName = new ReportModel
        {
            Person = new Person
            {
                Name = "Alice Johnson"
            }
        };
        engine.BuildReport(docWithName, modelWithName, "model");
        string reportWithNamePath = Path.Combine(workDir, "report_with_name.docx");
        docWithName.Save(reportWithNamePath);
    }
}
