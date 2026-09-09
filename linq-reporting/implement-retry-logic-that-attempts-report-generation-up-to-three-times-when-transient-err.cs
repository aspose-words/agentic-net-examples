using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Sample property used in the template.
    public string Name { get; set; } = "Aspose User";
}

public class Program
{
    public static void Main()
    {
        // Create a simple template document with a LINQ Reporting tag.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Report for <<[model.Name]>>");

        // Save the template to disk (required by the lifecycle rule).
        const string templatePath = "ReportTemplate.docx";
        template.Save(templatePath);

        // Load the template back (simulating a real scenario).
        Document doc = new Document(templatePath);

        // Prepare the data source.
        ReportModel model = new ReportModel();

        // Initialize the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Attempt to build the report up to three times in case of transient errors.
        const int maxAttempts = 3;
        int attempt = 0;
        bool success = false;

        while (attempt < maxAttempts && !success)
        {
            attempt++;
            try
            {
                // Build the report. The root object name must match the tag reference.
                success = engine.BuildReport(doc, model, "model");
                // If BuildReport returns false, treat it as a failure and retry.
                if (!success)
                {
                    Console.WriteLine($"Attempt {attempt}: BuildReport returned false.");
                }
            }
            catch (Exception ex)
            {
                // Log the transient error and continue to the next attempt.
                Console.WriteLine($"Attempt {attempt}: Transient error encountered - {ex.Message}");
            }

            if (!success && attempt < maxAttempts)
            {
                // Optional: wait briefly before retrying.
                System.Threading.Thread.Sleep(500);
            }
        }

        if (success)
        {
            // Save the generated report.
            const string outputPath = "GeneratedReport.docx";
            doc.Save(outputPath);
            Console.WriteLine($"Report generated successfully after {attempt} attempt(s). Saved to '{outputPath}'.");
        }
        else
        {
            Console.WriteLine($"Failed to generate report after {maxAttempts} attempts.");
        }
    }
}
