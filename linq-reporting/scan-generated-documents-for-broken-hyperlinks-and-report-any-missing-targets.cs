using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class HyperlinkScanner
{
    public static void Main()
    {
        // Ensure output directory exists.
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Create a sample file that will be linked correctly.
        string existingFilePath = Path.Combine(outputDir, "existing.txt");
        File.WriteAllText(existingFilePath, "This is an existing file.");

        // Path for the Word template.
        string templatePath = Path.Combine(outputDir, "template.docx");

        // Build a Word document with several hyperlinks.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Hyperlink scanning example:");
        builder.Writeln();

        // Hyperlink to an existing file (relative path).
        builder.InsertHyperlink("Existing File", "existing.txt", true);
        builder.Writeln();

        // Hyperlink to a missing file (relative path).
        builder.InsertHyperlink("Missing File", "missing.txt", true);
        builder.Writeln();

        // Hyperlink to a web URL.
        builder.InsertHyperlink("Web Link", "https://example.com", true);
        builder.Writeln();

        // Save the template document.
        doc.Save(templatePath);

        // Load the document for scanning.
        var loadedDoc = new Document(templatePath);

        // Scan for broken file hyperlinks.
        var brokenLinks = new List<string>();
        foreach (Field field in loadedDoc.Range.Fields)
        {
            if (field.Type == FieldType.FieldHyperlink)
            {
                var hyperlink = (FieldHyperlink)field;
                string address = hyperlink.Address;

                // Skip empty addresses.
                if (string.IsNullOrEmpty(address))
                    continue;

                // Treat HTTP/HTTPS links as valid without checking.
                if (address.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                    address.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                // Resolve the address relative to the document folder.
                string resolvedPath = Path.GetFullPath(Path.Combine(outputDir, address));

                // If the target file does not exist, record it.
                if (!File.Exists(resolvedPath))
                {
                    brokenLinks.Add(address);
                }
            }
        }

        // Report results.
        Console.WriteLine("Hyperlink scan completed.");
        if (brokenLinks.Count == 0)
        {
            Console.WriteLine("No broken file hyperlinks were found.");
        }
        else
        {
            Console.WriteLine("Broken file hyperlinks:");
            foreach (string link in brokenLinks)
            {
                Console.WriteLine($"- {link}");
            }
        }

        // Optionally, save a copy of the scanned document.
        string resultPath = Path.Combine(outputDir, "scanned.docx");
        loadedDoc.Save(resultPath);
    }
}
