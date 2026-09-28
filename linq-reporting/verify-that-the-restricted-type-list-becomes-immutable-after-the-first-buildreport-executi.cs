using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    public string Name { get; set; } = "Sample Name";
}

public class Program
{
    public static void Main()
    {
        // Prepare file paths in the current working directory.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");

        // -----------------------------------------------------------------
        // Create a simple template containing a LINQ Reporting tag.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("<<[model.Name]>>");
        template.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Sample data model.
        Model model = new Model();

        // -----------------------------------------------------------------
        // Configure the ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();

        // Use reflection to obtain the RestrictedTypes collection (the property
        // may not be publicly exposed in older library versions).
        PropertyInfo restrictedProp = typeof(ReportingEngine).GetProperty("RestrictedTypes",
            BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

        if (restrictedProp == null)
        {
            Console.WriteLine("RestrictedTypes property not found on ReportingEngine.");
            return;
        }

        // The collection is expected to implement IList<Type>.
        var restrictedTypes = (IList<Type>)restrictedProp.GetValue(engine)!;

        // Add a type to the restricted list before the first BuildReport.
        restrictedTypes.Add(typeof(Model));

        // Build the report.
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // Verify that the restricted type list becomes immutable after BuildReport.
        // -----------------------------------------------------------------
        bool isImmutable = false;
        try
        {
            // Attempt to modify the collection after BuildReport.
            restrictedTypes.Add(typeof(string));
        }
        catch (NotSupportedException)
        {
            // Expected: the collection is read‑only.
            isImmutable = true;
        }
        catch (Exception)
        {
            // Any other exception also indicates the list is not modifiable as expected.
            isImmutable = true;
        }

        Console.WriteLine($"RestrictedTypes immutable after BuildReport: {isImmutable}");

        // Save the generated report.
        reportDoc.Save(reportPath);
    }
}
