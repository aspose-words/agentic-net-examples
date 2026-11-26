using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required by Aspose.Words for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a simple template document with a LINQ Reporting tag.
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Report for: <<[model.Name]>>");
        templateDoc.Save(templatePath);

        // Define a simple data model.
        ReportModel model = new ReportModel { Name = "John Doe" };

        // Create the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Access the (possibly non‑public) RestrictedTypes collection via reflection.
        PropertyInfo restrictedProp = typeof(ReportingEngine).GetProperty(
            "RestrictedTypes",
            BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

        if (restrictedProp == null)
        {
            Console.WriteLine("Unable to locate RestrictedTypes property on ReportingEngine.");
            return;
        }

        // The collection is expected to implement IList<Type>.
        if (restrictedProp.GetValue(engine) is not IList<Type> restrictedTypes)
        {
            Console.WriteLine("RestrictedTypes property does not return IList<Type>.");
            return;
        }

        // Add the model type to the restricted list before building the report.
        restrictedTypes.Add(typeof(ReportModel));

        // Build the report for the first time.
        bool firstBuildSuccess = engine.BuildReport(templateDoc, model, "model");
        Console.WriteLine($"First BuildReport succeeded: {firstBuildSuccess}");

        // Save the generated report.
        string outputPath = "output.docx";
        templateDoc.Save(outputPath);
        Console.WriteLine($"Report saved to: {Path.GetFullPath(outputPath)}");

        // Verify that the restricted types list becomes immutable after the first BuildReport.
        try
        {
            // Attempt to modify the list after the report has been built.
            restrictedTypes.Add(typeof(string));
            Console.WriteLine("ERROR: RestrictedTypes list was modified after BuildReport – immutability test failed.");
        }
        catch (Exception ex)
        {
            // Expected path: an exception should be thrown.
            Console.WriteLine($"RestrictedTypes list is immutable after BuildReport. Caught exception: {ex.GetType().Name} - {ex.Message}");
        }
    }

    // Simple public data model class.
    public class ReportModel
    {
        public string Name { get; set; } = string.Empty;
    }
}
