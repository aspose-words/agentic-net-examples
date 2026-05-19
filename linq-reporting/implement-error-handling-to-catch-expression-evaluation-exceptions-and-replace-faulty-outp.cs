using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for template and result documents.
        const string templatePath = "template.docx";
        const string resultPath = "output.docx";

        // -----------------------------------------------------------------
        // 1. Create a template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template for reporting.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the data model.
        // -----------------------------------------------------------------
        var model = new ReportModel
        {
            person = new Person()
        };

        // -----------------------------------------------------------------
        // 4. Configure the ReportingEngine to inline error messages.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.InlineErrorMessages
        };

        // Build the report; any expression errors will be inlined.
        bool success;
        try
        {
            success = engine.BuildReport(doc, model, "model");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Unexpected error during report generation: " + ex.Message);
            success = false;
        }

        // -----------------------------------------------------------------
        // 5. Replace inline error messages with a placeholder text.
        // -----------------------------------------------------------------
        if (success)
        {
            // The engine inserts messages like "Error evaluating expression".
            // Replace them with "[Invalid]".
            doc.Range.Replace(new Regex(@"Error evaluating expression.*"), "[Invalid]");
        }

        // -----------------------------------------------------------------
        // 6. Save the final document.
        // -----------------------------------------------------------------
        doc.Save(resultPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class ReportModel
{
    // Initialized to avoid nullable warnings.
    public Person person { get; set; } = new Person();
}

public class Person
{
    // This property throws to simulate a faulty expression.
    public string Name
    {
        get => throw new InvalidOperationException("Simulated evaluation failure.");
    }

    // Normal property that will be rendered correctly.
    public int Age => 30;
}
