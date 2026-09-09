using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for all generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a dummy target file that will be linked correctly.
        string existingFilePath = Path.Combine(outputDir, "target.txt");
        File.WriteAllText(existingFilePath, "This is a valid target file.");

        // Path for a non‑existent file to simulate a broken link.
        string missingFilePath = Path.Combine(outputDir, "missing.txt");

        // Build a sample Word document containing both a valid and a broken hyperlink.
        string docPath = Path.Combine(outputDir, "Sample.docx");
        CreateSampleDocument(docPath, existingFilePath, missingFilePath);

        // Load the document and scan for broken hyperlinks.
        Document doc = new Document(docPath);
        List<string> brokenLinks = FindBrokenHyperlinks(doc, outputDir);

        // Report the results.
        Console.WriteLine("Hyperlink scan report:");
        if (brokenLinks.Count == 0)
        {
            Console.WriteLine("  No broken hyperlinks were found.");
        }
        else
        {
            foreach (string link in brokenLinks)
                Console.WriteLine($"  Broken link: {link}");
        }
    }

    // Creates a Word document with two hyperlinks: one valid, one broken.
    private static void CreateSampleDocument(string docPath, string validTarget, string invalidTarget)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Hyperlink scan example:");
        builder.Writeln();

        // Insert a hyperlink that points to an existing file.
        builder.Font.Color = System.Drawing.Color.Blue;
        builder.Font.Underline = Underline.Single;
        builder.InsertHyperlink("Valid Link", validTarget, false);
        builder.Writeln();

        // Insert a hyperlink that points to a missing file.
        builder.Font.Color = System.Drawing.Color.Blue;
        builder.Font.Underline = Underline.Single;
        builder.InsertHyperlink("Broken Link", invalidTarget, false);
        builder.Writeln();

        // Save the document.
        doc.Save(docPath);
    }

    // Scans the provided document for hyperlinks whose targets cannot be resolved.
    private static List<string> FindBrokenHyperlinks(Document doc, string baseDir)
    {
        var broken = new List<string>();

        foreach (Field field in doc.Range.Fields)
        {
            if (field.Type != FieldType.FieldHyperlink)
                continue;

            var hyperlink = (FieldHyperlink)field;
            string address = hyperlink.Address ?? string.Empty;

            // If the address is empty, consider it broken.
            if (string.IsNullOrWhiteSpace(address))
            {
                broken.Add("(empty address)");
                continue;
            }

            // Determine whether the address is a local file path.
            bool isLocalFile = !address.StartsWith("http://", StringComparison.OrdinalIgnoreCase) &&
                               !address.StartsWith("https://", StringComparison.OrdinalIgnoreCase);

            if (isLocalFile)
            {
                // Resolve relative paths against the document's folder.
                string resolvedPath = Path.IsPathRooted(address)
                    ? address
                    : Path.Combine(baseDir, address);

                if (!File.Exists(resolvedPath))
                    broken.Add(resolvedPath);
            }
            else
            {
                // For URLs we could attempt a network check, but to keep the example self‑contained,
                // we treat all URLs as valid.
            }
        }

        return broken;
    }
}
