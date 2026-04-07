using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class HyperlinkScanner
{
    public static void Main()
    {
        // Prepare a working directory.
        string workDir = Directory.GetCurrentDirectory();

        // Create a dummy file that will be linked correctly.
        string existingFileName = "ExistingFile.txt";
        string existingFilePath = Path.Combine(workDir, existingFileName);
        File.WriteAllText(existingFilePath, "This file exists.");

        // Define the name of a file that will NOT exist.
        string missingFileName = "MissingFile.txt";

        // Create a new Word document and add hyperlinks.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Hyperlink to an existing local file.
        builder.InsertHyperlink("Existing File", existingFileName, false);
        builder.Writeln();

        // Hyperlink to a missing local file.
        builder.InsertHyperlink("Missing File", missingFileName, false);
        builder.Writeln();

        // Hyperlink to an external URL (treated as valid).
        builder.InsertHyperlink("Web Site", "https://www.example.com", false);
        builder.Writeln();

        // Save the generated document.
        string docPath = Path.Combine(workDir, "GeneratedDoc.docx");
        doc.Save(docPath);

        // Load the document for scanning.
        Document loadedDoc = new Document(docPath);

        // Scan all hyperlink fields.
        List<string> brokenLinks = new List<string>();
        foreach (Field field in loadedDoc.Range.Fields)
        {
            // Only process hyperlink fields.
            if (field.Type != FieldType.FieldHyperlink)
                continue;

            FieldHyperlink hyperlink = (FieldHyperlink)field;
            string address = hyperlink.Address;

            if (string.IsNullOrEmpty(address))
                continue;

            // Skip HTTP/HTTPS links – assume they are reachable.
            if (address.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                continue;

            // Resolve relative paths against the document folder.
            string targetPath = Path.IsPathRooted(address)
                ? address
                : Path.Combine(workDir, address);

            if (!File.Exists(targetPath))
                brokenLinks.Add(address);
        }

        // Report the results.
        if (brokenLinks.Count == 0)
        {
            Console.WriteLine("No broken hyperlinks were found.");
        }
        else
        {
            Console.WriteLine("Broken hyperlinks detected:");
            foreach (string link in brokenLinks)
                Console.WriteLine($" - {link}");
        }
    }
}
