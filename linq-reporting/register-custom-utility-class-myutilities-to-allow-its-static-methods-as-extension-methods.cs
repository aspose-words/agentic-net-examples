using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public static class MyUtilities
{
    // Extension-like static method that can be called from LINQ Reporting tags.
    public static string ToUpper(string value) => value?.ToUpperInvariant() ?? string.Empty;
}

public class Person
{
    // Initialize to avoid nullable warnings.
    public string Name { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a template document with a LINQ Reporting tag that calls
        //    the static utility method.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Correct tag syntax for calling a static method: <<[MyUtilities.ToUpper(model.Name)]>>
        builder.Writeln("<<[MyUtilities.ToUpper(model.Name)]>>");

        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template for report generation.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the data model.
        // -----------------------------------------------------------------
        var model = new Person { Name = "John Doe" };

        // -----------------------------------------------------------------
        // 4. Configure the reporting engine and register the utility class.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine();
        engine.KnownTypes.Add(typeof(MyUtilities));

        // -----------------------------------------------------------------
        // 5. Build the report using the model and the root name "model".
        // -----------------------------------------------------------------
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 6. Save the generated report.
        // -----------------------------------------------------------------
        doc.Save("Report.docx");
    }
}
