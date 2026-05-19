using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Folder for temporary files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // Create a sample document with several hyperlinks.
        string docPath = Path.Combine(workDir, "Sample.docx");
        CreateSampleDocument(docPath, workDir);

        // Scan the document for broken hyperlinks.
        List<string> brokenLinks = ScanDocumentForBrokenHyperlinks(docPath, workDir);

        // Output the report.
        Console.WriteLine("Broken hyperlink report:");
        if (brokenLinks.Count == 0)
        {
            Console.WriteLine("No broken hyperlinks were found.");
        }
        else
        {
            foreach (string entry in brokenLinks)
                Console.WriteLine(entry);
        }

        // Save the report to a text file.
        string reportPath = Path.Combine(workDir, "BrokenLinksReport.txt");
        File.WriteAllLines(reportPath, brokenLinks);
    }

    // Creates a document containing a mix of valid and invalid hyperlinks.
    private static void CreateSampleDocument(string docPath, string workDir)
    {
        // Create a dummy file that will be linked correctly.
        string existingFile = Path.Combine(workDir, "ExistingFile.txt");
        File.WriteAllText(existingFile, "This file exists.");

        // Define a non‑existent file path.
        string missingFile = Path.Combine(workDir, "MissingFile.txt");

        // Define a URL (will not be validated in this example).
        string url = "https://example.com";

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Valid local file hyperlink.
        builder.InsertHyperlink("Valid File Link", existingFile, false);
        builder.Writeln();

        // Invalid local file hyperlink.
        builder.InsertHyperlink("Broken File Link", missingFile, false);
        builder.Writeln();

        // URL hyperlink (treated as external, not checked for existence).
        builder.InsertHyperlink("Web Link", url, false);
        builder.Writeln();

        // Save the document.
        doc.Save(docPath);
    }

    // Scans a document for hyperlinks whose targets cannot be resolved.
    private static List<string> ScanDocumentForBrokenHyperlinks(string docPath, string workDir)
    {
        List<string> broken = new List<string>();

        Document doc = new Document(docPath);

        // Iterate over all fields and pick out hyperlink fields.
        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type != FieldType.FieldHyperlink)
                continue;

            FieldHyperlink hyperlink = (FieldHyperlink)field;
            string address = hyperlink.Address?.Trim() ?? string.Empty;
            string display = hyperlink.DisplayResult?.Trim() ?? string.Empty;

            // Skip empty addresses.
            if (string.IsNullOrEmpty(address))
                continue;

            // Determine if the address is a URL.
            bool isUrl = address.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                         address.StartsWith("https://", StringComparison.OrdinalIgnoreCase);

            if (isUrl)
            {
                // For this example we do not attempt to reach the URL.
                // You could add an HTTP request here if desired.
                continue;
            }

            // Resolve relative paths against the document's folder.
            string targetPath = address;
            if (!Path.IsPathRooted(targetPath))
                targetPath = Path.Combine(workDir, targetPath);

            if (!File.Exists(targetPath))
            {
                broken.Add($"Display Text: \"{display}\", Target: \"{address}\" (File not found)");
            }
        }

        return broken;
    }
}
