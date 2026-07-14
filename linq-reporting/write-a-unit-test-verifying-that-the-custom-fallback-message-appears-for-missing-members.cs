using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // 1. Create a template document with a tag that references a missing member.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        // The tag <<[missingObject.Name]>> refers to a member that does not exist.
        builder.Writeln("<<[missingObject.Name]>>");

        // Save the template to disk.
        string templatePath = Path.Combine(outputDir, "template.docx");
        template.Save(templatePath);

        // 2. Load the template back (required by the lifecycle rule).
        Document loadedTemplate = new Document(templatePath);

        // 3. Prepare a dummy data source (empty DataSet is sufficient).
        DataSet dummyData = new DataSet();

        // 4. Configure the ReportingEngine to allow missing members and set a custom fallback message.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.AllowMissingMembers;
        engine.MissingMemberMessage = "CustomMissingMessage";

        // 5. Build the report. The root name is irrelevant because we only use a plain reference.
        bool buildResult = engine.BuildReport(loadedTemplate, dummyData, "");

        // 6. Save the generated report.
        string resultPath = Path.Combine(outputDir, "result.docx");
        loadedTemplate.Save(resultPath);

        // 7. Verify that the custom fallback message appears in the output document text.
        string resultText = loadedTemplate.GetText();

        if (resultText.Contains(engine.MissingMemberMessage))
        {
            Console.WriteLine("Test passed: custom fallback message was inserted.");
        }
        else
        {
            Console.WriteLine("Test failed: custom fallback message was not found.");
        }

        // Optional: output the paths for reference.
        Console.WriteLine($"Template saved to: {templatePath}");
        Console.WriteLine($"Result saved to: {resultPath}");
    }
}
