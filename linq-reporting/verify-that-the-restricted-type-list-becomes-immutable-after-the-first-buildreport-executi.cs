using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    public string Name { get; set; } = "Aspose";
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a simple template with a LINQ Reporting tag.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);
        builder.Writeln("<<[model.Name]>>");
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template for reporting.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Define restricted types BEFORE the first BuildReport call.
        // -----------------------------------------------------------------
        ReportingEngine.SetRestrictedTypes(typeof(Environment));

        // -----------------------------------------------------------------
        // 4. Build the first report.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(doc, new Model(), "model");

        // -----------------------------------------------------------------
        // 5. Verify that the restricted type list is now immutable.
        // -----------------------------------------------------------------
        bool isImmutable = false;
        try
        {
            // Attempt to modify the restricted types after BuildReport.
            ReportingEngine.SetRestrictedTypes(typeof(System.IO.File));
        }
        catch (InvalidOperationException)
        {
            // Expected exception indicates immutability.
            isImmutable = true;
        }

        // -----------------------------------------------------------------
        // 6. Output verification result.
        // -----------------------------------------------------------------
        Console.WriteLine(isImmutable
            ? "Restricted type list is immutable after first BuildReport."
            : "Restricted type list is still mutable (unexpected).");

        // -----------------------------------------------------------------
        // 7. Save the generated report.
        // -----------------------------------------------------------------
        const string reportPath = "Report.docx";
        doc.Save(reportPath);
    }
}
