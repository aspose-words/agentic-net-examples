using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Path for the temporary template document.
        const string templatePath = "template.docx";

        // -------------------------------------------------
        // 1. Create a template that accesses a restricted type.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // The template creates a variable of type System.Type and then tries to read its FullName.
        // Access to System.Type will be restricted later.
        builder.Writeln("<<var [typeVar = typeof(string)]>><<[typeVar.FullName]>>");

        // Save the template to disk before building the report (required by the lifecycle rule).
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the saved template.
        // -------------------------------------------------
        Document doc = new Document(templatePath);

        // -------------------------------------------------
        // 3. Set restricted types BEFORE the first BuildReport call.
        // -------------------------------------------------
        ReportingEngine.SetRestrictedTypes(typeof(System.Type));

        // -------------------------------------------------
        // 4. Build the report and verify that an exception is thrown.
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine();

        try
        {
            // No data source is needed for this template; an empty object is sufficient.
            engine.BuildReport(doc, new object());

            // If we reach this line, the engine did NOT enforce the restriction as expected.
            Console.WriteLine("Test FAILED: No exception was thrown.");
        }
        catch (Exception ex)
        {
            // The engine should throw an exception because access to System.Type is prohibited.
            Console.WriteLine($"Test PASSED: Caught expected exception -> {ex.GetType().Name}");
        }
    }
}
