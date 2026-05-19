using System;
using System.Data;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare folders.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a template document with tags that reference missing members.
        // -----------------------------------------------------------------
        DocumentBuilder builder = new DocumentBuilder();
        builder.Writeln("Customer Name: <<[customer.Name]>>");
        builder.Writeln("<<foreach [order in orders]>>Order ID: <<[order.Id]>> <</foreach>>");
        string templatePath = Path.Combine(outputDir, "template.docx");
        builder.Document.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template back for reporting.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Configure the ReportingEngine to allow missing members.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.AllowMissingMembers;
        engine.MissingMemberMessage = "MISSING_MEMBER";

        // Use an empty DataSet as the data source – it contains no members referenced in the template.
        DataSet emptyData = new DataSet();

        // Build the report. The root name is empty because we do not reference the data source object itself.
        engine.BuildReport(doc, emptyData, "");

        // -----------------------------------------------------------------
        // 4. Log each occurrence of the missing member placeholder.
        // -----------------------------------------------------------------
        string documentText = doc.GetText();
        int missingCount = Regex.Matches(documentText, Regex.Escape(engine.MissingMemberMessage)).Count;

        string logMessage = $"Missing member occurrences: {missingCount}";
        Console.WriteLine(logMessage);

        string logPath = Path.Combine(outputDir, "missing_log.txt");
        File.WriteAllText(logPath, logMessage);

        // Save the final report.
        string resultPath = Path.Combine(outputDir, "result.docx");
        doc.Save(resultPath);
    }
}
