using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class SensitiveInfo
{
    public string Secret { get; set; } = string.Empty;
}

public class ReportModel
{
    public SensitiveInfo Sensitive { get; set; } = new SensitiveInfo();
}

public class Program
{
    public static void Main()
    {
        // Paths for the temporary template file.
        const string templatePath = "template.docx";

        // 1. Create a template document with a tag that accesses a member of the restricted type.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        // The tag tries to read Sensitive.Secret – Sensitive is of type SensitiveInfo.
        builder.Writeln("<<[model.Sensitive.Secret]>>");
        // Save the template to disk.
        templateDoc.Save(templatePath);

        // 2. Load the template back (simulating a real-world scenario).
        Document loadedDoc = new Document(templatePath);

        // 3. Define the restricted type before any report generation.
        ReportingEngine.SetRestrictedTypes(typeof(SensitiveInfo));

        // 4. Prepare the data model.
        ReportModel model = new ReportModel
        {
            Sensitive = new SensitiveInfo { Secret = "TopSecret" }
        };

        // 5. Build the report and verify that an exception is thrown because the template
        //    attempts to access a member of a restricted type.
        ReportingEngine engine = new ReportingEngine();

        bool exceptionThrown = false;
        try
        {
            // The root name used in the template is "model".
            engine.BuildReport(loadedDoc, model, "model");
        }
        catch (Exception ex)
        {
            // Expected path – the engine should reject access to the restricted type.
            exceptionThrown = true;
            Console.WriteLine($"Expected exception caught: {ex.GetType().Name} - {ex.Message}");
        }

        // 6. Report the test outcome.
        if (exceptionThrown)
        {
            Console.WriteLine("Test passed: Restricted type enforcement threw an exception as expected.");
        }
        else
        {
            Console.WriteLine("Test failed: No exception was thrown when accessing a restricted type.");
        }
    }
}
