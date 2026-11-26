using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    // Existing property (not used in the template).
    public string ExistingProperty { get; set; } = "ExistingValue";
}

public class Program
{
    public static void Main()
    {
        // 1. Create a template document that references a missing member.
        var template = new Document();
        var builder = new DocumentBuilder(template);
        builder.Writeln("Report start");
        builder.Writeln("Missing member: <<[model.MissingProperty]>>");
        builder.Writeln("Report end");

        // Save the template to disk – required by the workflow.
        const string templatePath = "template.docx";
        template.Save(templatePath);

        // 2. Load the template for reporting.
        var doc = new Document(templatePath);

        // 3. Prepare the root data object.
        var model = new Model();

        // 4. Configure the reporting engine with inline error messages.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // 5. Build the report. The method returns false because a missing
        //    member is encountered, but an inline error message should be inserted.
        bool success = engine.BuildReport(doc, model, "model");

        // 6. Verify that the inline error message appears in the generated document.
        //    The default inline error contains the missing member name.
        string resultText = doc.GetText();
        bool containsMissingMember = resultText.Contains("MissingProperty");

        // Simple assertion logic to emulate a unit test.
        if (!success && containsMissingMember)
        {
            Console.WriteLine("Test passed: inline error message displayed for missing member.");
        }
        else
        {
            Console.WriteLine("Test failed: inline error message not displayed as expected.");
        }

        // Save the generated report for manual inspection if needed.
        doc.Save("output.docx");
    }
}
