using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    // Initialize to avoid nullable warnings.
    public string Name { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "template.docx";
        const string outputPath = "report.docx";

        // -----------------------------------------------------------------
        // 1. Create a simple template document with a LINQ Reporting tag.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        // Tag references the root object named "model".
        builder.Writeln("Hello <<[model.Name]>>!");
        // Save the template so it exists on disk (optional for this example).
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Define restricted types BEFORE the first BuildReport call.
        // -----------------------------------------------------------------
        // For demonstration we restrict the System.String type.
        ReportingEngine.SetRestrictedTypes(typeof(string));

        // -----------------------------------------------------------------
        // 3. Build the report using the template and a data model.
        // -----------------------------------------------------------------
        Model data = new Model { Name = "World" };
        ReportingEngine engine = new ReportingEngine();
        // BuildReport overload that allows referencing the root object name.
        engine.BuildReport(templateDoc, data, "model");

        // -----------------------------------------------------------------
        // 4. Verify that the restricted type list is now immutable.
        //    Attempting to modify it should throw an exception.
        // -----------------------------------------------------------------
        bool isImmutable = false;
        try
        {
            // This call should fail because the restricted types have already been locked.
            ReportingEngine.SetRestrictedTypes(typeof(int));
        }
        catch (InvalidOperationException)
        {
            // Expected: the list cannot be changed after BuildReport.
            isImmutable = true;
        }
        catch (ArgumentException)
        {
            // In case the API throws ArgumentException for other validation errors.
            isImmutable = true;
        }

        // Output the verification result.
        Console.WriteLine($"Restricted types immutable after first BuildReport: {isImmutable}");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        templateDoc.Save(outputPath);
    }
}
