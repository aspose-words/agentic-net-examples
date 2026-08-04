using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    // Sample property used in the template.
    public string Name { get; set; } = "Aspose";
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string resultPath = "Result.docx";

        // -----------------------------------------------------------------
        // 1. Create a simple template document containing a LINQ Reporting tag.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("<<[model.Name]>>"); // Tag that will be replaced by Model.Name.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the first report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, new Model(), "model");

        // -----------------------------------------------------------------
        // 3. Attempt to set restricted types after the first BuildReport.
        //    According to the documentation this must throw an exception.
        // -----------------------------------------------------------------
        bool exceptionThrown = false;
        try
        {
            // Any public type can be passed; using System.String for simplicity.
            ReportingEngine.SetRestrictedTypes(typeof(string));
        }
        catch (InvalidOperationException)
        {
            // Expected exception when modifying restricted types after building a report.
            exceptionThrown = true;
        }

        // Output the test result.
        Console.WriteLine($"SetRestrictedTypes after BuildReport threw exception: {exceptionThrown}");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        reportDoc.Save(resultPath);
    }
}
