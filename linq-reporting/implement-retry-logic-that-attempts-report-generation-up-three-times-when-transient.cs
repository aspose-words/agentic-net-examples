using System;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Reporting;

public class ReportGenerator
{
    // Generates a report from a template with retry logic (max 3 attempts).
    public void GenerateReport(string templatePath, object dataSource, string outputPath)
    {
        const int maxAttempts = 3;
        int attempt = 0;
        bool success = false;

        // LoadOptions can be customized; here we keep the default recovery mode.
        LoadOptions loadOptions = new LoadOptions();

        while (attempt < maxAttempts && !success)
        {
            attempt++;
            try
            {
                // ----- Load the template document -----
                Document template = new Document(templatePath, loadOptions);

                // ----- Build the report -----
                ReportingEngine engine = new ReportingEngine();
                // You can set options if needed, e.g., engine.Options = ReportBuildOptions.AllowMissingMembers;
                engine.BuildReport(template, dataSource);

                // ----- Save the generated report -----
                template.Save(outputPath);

                success = true; // If we reach here, the operation succeeded.
            }
            catch (Exception ex) when (IsTransient(ex))
            {
                // Transient error encountered. If we have attempts left, retry.
                if (attempt >= maxAttempts)
                {
                    // No more retries; rethrow the exception.
                    throw new InvalidOperationException(
                        $"Report generation failed after {maxAttempts} attempts due to a transient error.", ex);
                }

                // Optionally, introduce a small delay before retrying.
                Thread.Sleep(500);
            }
        }
    }

    // Determines whether an exception is considered transient.
    private bool IsTransient(Exception ex)
    {
        // For demonstration, treat the following as transient:
        // - FileCorruptedException (document loading issues)
        // - IOException (I/O problems)
        // - Any other exception can be added as needed.
        return ex is FileCorruptedException ||
               ex is IOException;
    }
}

public static class Program
{
    // Entry point required for a console application.
    public static void Main(string[] args)
    {
        // Simple demonstration of the ReportGenerator usage.
        // Adjust the paths to point to real files when running the example.
        string templatePath = "Template.docx"; // path to the Aspose.Words template
        string outputPath   = "Report_Output.docx"; // where the generated report will be saved
        var dataSource = new { Name = "John Doe", Date = DateTime.Today }; // example anonymous data source

        var generator = new ReportGenerator();
        try
        {
            generator.GenerateReport(templatePath, dataSource, outputPath);
            Console.WriteLine($"Report generated successfully at '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Failed to generate report: {ex.Message}");
        }
    }
}
