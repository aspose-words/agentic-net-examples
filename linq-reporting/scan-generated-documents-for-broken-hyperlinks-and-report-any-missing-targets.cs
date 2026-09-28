using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Ensure output directory exists.
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Paths for template and generated documents.
        string templatePath = Path.Combine(outputDir, "template.docx");
        string resultPath = Path.Combine(outputDir, "result.docx");

        // -----------------------------------------------------------------
        // 1. Create a LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Links Report:");
        builder.Writeln("<<foreach [link in Links]>>");
        // Insert a hyperlink using LINQ Reporting tag.
        builder.Writeln("<<link [link.Url] [link.Text]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Prepare sample data model with one valid and one broken link.
        // -----------------------------------------------------------------
        // Create a file that will be referenced by a valid hyperlink.
        string existingFilePath = Path.Combine(outputDir, "existing.txt");
        File.WriteAllText(existingFilePath, "This is an existing file.");

        // Define the data model.
        var model = new ReportModel
        {
            Links = new List<LinkInfo>
            {
                new LinkInfo
                {
                    Url = existingFilePath,
                    Text = "Existing File"
                },
                new LinkInfo
                {
                    Url = Path.Combine(outputDir, "missing.txt"), // This file does not exist.
                    Text = "Missing File"
                }
            }
        };

        // -----------------------------------------------------------------
        // 3. Build the report using Aspose.Words ReportingEngine.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");
        reportDoc.Save(resultPath);

        // -----------------------------------------------------------------
        // 4. Scan the generated document for broken hyperlinks.
        // -----------------------------------------------------------------
        Document generatedDoc = new Document(resultPath);
        List<string> brokenLinks = new List<string>();

        foreach (FieldHyperlink hyperlink in generatedDoc.Range.Fields.OfType<FieldHyperlink>())
        {
            string address = hyperlink.Address;

            // If the address is a file path, verify its existence.
            // For simplicity, treat non‑file URLs as valid.
            if (!string.IsNullOrEmpty(address) && !address.StartsWith("http", StringComparison.OrdinalIgnoreCase))
            {
                // Normalize possible "file://" prefix.
                string normalizedPath = address;
                if (address.StartsWith("file://", StringComparison.OrdinalIgnoreCase))
                {
                    normalizedPath = new Uri(address).LocalPath;
                }

                if (!File.Exists(normalizedPath))
                {
                    brokenLinks.Add(address);
                }
            }
        }

        // -----------------------------------------------------------------
        // 5. Output the scan results.
        // -----------------------------------------------------------------
        Console.WriteLine("Hyperlink scan completed.");
        if (brokenLinks.Count == 0)
        {
            Console.WriteLine("No broken hyperlinks were found.");
        }
        else
        {
            Console.WriteLine("Broken hyperlinks:");
            foreach (string link in brokenLinks)
            {
                Console.WriteLine($"- {link}");
            }
        }
    }
}

// ---------------------------------------------------------------------
// Data model classes used by the LINQ Reporting engine.
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<LinkInfo> Links { get; set; } = new();
}

public class LinkInfo
{
    public string Url { get; set; } = string.Empty;
    public string Text { get; set; } = string.Empty;
}
