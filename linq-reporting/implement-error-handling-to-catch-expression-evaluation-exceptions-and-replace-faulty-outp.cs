using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Replacing;

public class Order
{
    public int Value { get; set; } = 10;
    public int Divisor { get; set; } = 0; // Will cause divide‑by‑zero in the template
}

public class Program
{
    public static void Main()
    {
        // Prepare working directory and file paths
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "template.docx");
        string outputPath = Path.Combine(workDir, "output.docx");

        // -----------------------------------------------------------------
        // 1. Create the template document with a faulty expression tag.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Result of division: <<[order.Value / order.Divisor]>>");
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template (simulating a separate load step).
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the data model.
        // -----------------------------------------------------------------
        Order order = new Order();

        // -----------------------------------------------------------------
        // 4. Configure the reporting engine to embed inline error messages.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // -----------------------------------------------------------------
        // 5. Build the report with error handling.
        //    If an exception occurs, replace the faulty tag with a placeholder.
        // -----------------------------------------------------------------
        bool success = false;
        try
        {
            success = engine.BuildReport(doc, order, "order");
        }
        catch (Exception ex)
        {
            // Replace the problematic expression with a placeholder text.
            FindReplaceOptions replaceOptions = new FindReplaceOptions
            {
                MatchCase = false,
                FindWholeWordsOnly = true
            };
            doc.Range.Replace("<<[order.Value / order.Divisor]>>", "N/A", replaceOptions);
            Console.WriteLine($"Error during report generation: {ex.Message}");
        }

        // -----------------------------------------------------------------
        // 6. Save the generated report.
        // -----------------------------------------------------------------
        doc.Save(outputPath);

        // -----------------------------------------------------------------
        // 7. Inform the user about the result.
        // -----------------------------------------------------------------
        Console.WriteLine($"Report build success: {success}");
        Console.WriteLine($"Output saved to: {outputPath}");
    }
}
