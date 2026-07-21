using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the working directory exists.
        const string templatePath = "Template.docx";
        const string outputPath = "ReportOutput.docx";

        // Step 1: Create a LINQ Reporting template programmatically.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Report Title: <<[model.Title]>>");
        builder.Writeln("<<foreach [item in model.Items]>>");
        builder.Writeln("- <<[item.Name]>>: <<[item.Value]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Step 2: Load the template for report generation.
        var doc = new Document(templatePath);

        // Step 3: Prepare the data model.
        var model = new ReportModel
        {
            Title = "Sales Summary",
            Items = new List<ReportItem>
            {
                new ReportItem { Name = "Product A", Value = 1200 },
                new ReportItem { Name = "Product B", Value = 850 },
                new ReportItem { Name = "Product C", Value = 430 }
            }
        };

        // Step 4: Build the report with retry logic (up to 3 attempts).
        var engine = new ReportingEngine();
        bool success = false;
        const int maxAttempts = 3;

        for (int attempt = 1; attempt <= maxAttempts; attempt++)
        {
            try
            {
                // BuildReport overload that includes the root object name.
                engine.BuildReport(doc, model, "model");
                success = true;
                break; // Exit loop on success.
            }
            catch (Exception ex) when (IsTransient(ex) && attempt < maxAttempts)
            {
                // Transient error encountered – wait briefly before retrying.
                Console.WriteLine($"Attempt {attempt} failed with a transient error: {ex.Message}");
                Thread.Sleep(500); // Simple back‑off delay.
            }
        }

        if (!success)
        {
            Console.WriteLine("Report generation failed after multiple attempts.");
            return;
        }

        // Step 5: Save the generated report.
        doc.Save(outputPath);
        Console.WriteLine($"Report generated successfully: {Path.GetFullPath(outputPath)}");
    }

    // Simple heuristic to decide whether an exception is transient.
    private static bool IsTransient(Exception ex)
    {
        // In a real scenario, inspect the exception type/message.
        // For this example, treat all exceptions as transient.
        return true;
    }
}

// Data model for the report.
public class ReportModel
{
    public string Title { get; set; } = string.Empty;
    public List<ReportItem> Items { get; set; } = new();
}

// Individual item displayed in the report.
public class ReportItem
{
    public string Name { get; set; } = string.Empty;
    public int Value { get; set; }
}
