using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public string CustomerName { get; set; } = "John Doe";
    public List<string> Items { get; set; } = new() { "Item A", "Item B", "Item C" };
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for any legacy encodings Aspose.Words might need.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for the template and the final report.
        const string templatePath = "template.docx";
        const string reportPath = "report.docx";

        // -----------------------------------------------------------------
        // Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Simple report showing a customer name and a list of items.
        builder.Writeln("Report for <<[order.CustomerName]>>");
        builder.Writeln("<<foreach [item in order.Items]>>");
        builder.Writeln("- <<[item]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template back for report generation.
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // Prepare the data source.
        Order order = new Order();

        // -----------------------------------------------------------------
        // Retry logic: attempt to build the report up to three times if a
        // transient exception occurs.
        // -----------------------------------------------------------------
        const int maxAttempts = 3;
        int attempt = 0;
        bool success = false;

        while (attempt < maxAttempts && !success)
        {
            attempt++;
            try
            {
                ReportingEngine engine = new ReportingEngine();
                // BuildReport returns a bool only when InlineErrorMessages is set; we ignore the return value here.
                engine.BuildReport(loadedTemplate, order, "order");
                success = true;
            }
            catch (IOException ex) // Example of a transient error type.
            {
                Console.WriteLine($"Attempt {attempt} failed with transient error: {ex.Message}");
                if (attempt == maxAttempts)
                {
                    Console.WriteLine("Maximum retry attempts reached. Rethrowing exception.");
                    throw;
                }
                // Optionally add a short delay before retrying.
                System.Threading.Thread.Sleep(500);
            }
            catch (Exception ex)
            {
                // Non-transient errors are not retried.
                Console.WriteLine($"Non-transient error encountered: {ex.Message}");
                throw;
            }
        }

        // -----------------------------------------------------------------
        // Save the generated report.
        // -----------------------------------------------------------------
        if (success)
        {
            loadedTemplate.Save(reportPath);
            Console.WriteLine($"Report generated successfully and saved to '{reportPath}'.");
        }
    }
}
