using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Folder for temporary files.
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "template.docx");
        string resultPath = Path.Combine(workDir, "result.docx");

        // 1. Create a template document with a tag that references a missing member.
        DocumentBuilder builder = new DocumentBuilder();
        // The tag <<[missingObject.Name]>> refers to a member that does not exist.
        builder.Writeln("<<[missingObject.Name]>>");
        // Save the template so that the engine can load it later.
        builder.Document.Save(templatePath);

        // 2. Load the template document.
        Document doc = new Document(templatePath);

        // 3. Configure the ReportingEngine to allow missing members and set a custom fallback message.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.AllowMissingMembers;
        string fallbackMessage = "CustomMissing";
        engine.MissingMemberMessage = fallbackMessage;

        // 4. Build the report. The data source can be any object because we are only testing missing members.
        // The empty string for the data source name means we will reference members directly.
        engine.BuildReport(doc, new object(), "");

        // 5. Save the generated document.
        doc.Save(resultPath);

        // 6. Verify that the custom fallback message appears in the output.
        string resultText = doc.GetText();
        if (resultText.Contains(fallbackMessage))
        {
            Console.WriteLine("Test passed: custom fallback message was inserted.");
        }
        else
        {
            Console.WriteLine("Test failed: custom fallback message was not found.");
        }
    }
}
