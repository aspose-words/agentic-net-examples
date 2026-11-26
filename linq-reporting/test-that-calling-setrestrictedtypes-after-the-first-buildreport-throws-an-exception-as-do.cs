using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    public string Name { get; set; } = "World";
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // Create a simple template with a LINQ Reporting tag.
        // -----------------------------------------------------------------
        const string templatePath = "template.docx";

        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);
        builder.Writeln("Hello <<[model.Name]>>");
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template for reporting.
        // -----------------------------------------------------------------
        Document doc = new(templatePath);
        Model model = new();

        // -----------------------------------------------------------------
        // Set restricted types BEFORE the first BuildReport.
        // SetRestrictedTypes is a static method, so call it on the type.
        // -----------------------------------------------------------------
        ReportingEngine.SetRestrictedTypes(typeof(string));

        // Build the report – this should succeed.
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // -----------------------------------------------------------------
        // Attempt to call SetRestrictedTypes AFTER the first BuildReport.
        // This must throw InvalidOperationException as documented.
        // -----------------------------------------------------------------
        try
        {
            ReportingEngine.SetRestrictedTypes(typeof(int));
            Console.WriteLine("Error: No exception was thrown when calling SetRestrictedTypes after BuildReport.");
        }
        catch (InvalidOperationException ex)
        {
            Console.WriteLine($"Expected exception caught: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected exception type: {ex.GetType().Name} - {ex.Message}");
        }
    }
}
