using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data model.
        var model = new ReportModel
        {
            Name = "Sample Report"
        };

        // Create a template document with a tag that attempts to write a file.
        var templatePath = "template.docx";
        CreateTemplate(templatePath);

        // Load the template.
        var doc = new Document(templatePath);

        // Configure ReportingEngine.
        var engine = new ReportingEngine();

        // Attempt to configure restricted members via reflection (to stay compatible with different library versions).
        ConfigureRestrictedMembers(engine);

        // Build the report and capture any security exception.
        bool success;
        try
        {
            success = engine.BuildReport(doc, model, "model");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Report generation failed: {ex.Message}");
            success = false;
        }

        // Save the result if generation succeeded.
        if (success)
        {
            var outputPath = "result.docx";
            doc.Save(outputPath);
            Console.WriteLine($"Report generated successfully: {outputPath}");
        }
        else
        {
            Console.WriteLine("Report generation was blocked due to restricted members.");
        }
    }

    private static void CreateTemplate(string path)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Simple title.
        builder.Writeln("<<[model.Name]>>");
        builder.Writeln();

        // Attempt to call a method that writes a file (should be blocked).
        builder.Writeln("Attempting file write:");
        builder.Writeln("<<[System.IO.File.WriteAllText(\"blocked.txt\", \"This should be blocked.\")]>>");

        // Save the template.
        doc.Save(path);
    }

    private static void ConfigureRestrictedMembers(ReportingEngine engine)
    {
        // Use reflection to avoid compile‑time dependency on the exact API version.
        var engineType = engine.GetType();
        var restrictedProp = engineType.GetProperty("RestrictedMembers");
        if (restrictedProp == null) return;

        if (restrictedProp.GetValue(engine) is IList list)
        {
            // Find the RestrictedMemberInfo type.
            var rmiType = engineType.Assembly.GetType("Aspose.Words.Reporting.RestrictedMemberInfo");
            if (rmiType == null) return;

            var createMethod = rmiType.GetMethod("CreateMethod", new[] { typeof(Type), typeof(string) });
            if (createMethod == null) return;

            var dangerousMethods = new[]
            {
                "WriteAllText",
                "WriteAllBytes",
                "AppendAllText",
                "Create",
                "Open",
                "OpenWrite"
            };

            foreach (var methodName in dangerousMethods)
            {
                var rmi = createMethod.Invoke(null, new object[] { typeof(File), methodName });
                list.Add(rmi);
            }
        }
    }
}

public class ReportModel
{
    public string Name { get; set; } = string.Empty;
}
