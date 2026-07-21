using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

namespace HyperlinkScanner
{
    public class Program
    {
        public static void Main()
        {
            // Define file names in the current working directory.
            string docPath = Path.Combine(Directory.GetCurrentDirectory(), "Generated.docx");
            string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "BrokenLinksReport.txt");

            // -----------------------------------------------------------------
            // Step 1: Create a sample document containing several hyperlinks.
            // -----------------------------------------------------------------
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Ensure the target folder for a valid file exists.
            string existingFileName = "ExistingFile.txt";
            File.WriteAllText(existingFileName, "Sample content"); // create a real file.

            // Insert a hyperlink that points to an existing local file.
            builder.InsertHyperlink("Existing File", existingFileName, false);
            builder.Writeln();

            // Insert a hyperlink that points to a missing local file.
            builder.InsertHyperlink("Missing File", "MissingFile.txt", false);
            builder.Writeln();

            // Insert a hyperlink that points to a web URL (treated as valid for this example).
            builder.InsertHyperlink("Web Link", "https://www.example.com", false);
            builder.Writeln();

            // Save the generated document.
            doc.Save(docPath);

            // -----------------------------------------------------------------
            // Step 2: Load the document and scan for broken hyperlinks.
            // -----------------------------------------------------------------
            Document loadedDoc = new Document(docPath);
            List<string> brokenLinks = new List<string>();

            foreach (Field field in loadedDoc.Range.Fields)
            {
                if (field.Type != FieldType.FieldHyperlink)
                    continue;

                FieldHyperlink hyperlink = (FieldHyperlink)field;
                string address = hyperlink.Address ?? string.Empty;

                // Skip empty addresses and bookmark links.
                if (string.IsNullOrWhiteSpace(address))
                    continue;

                // Treat HTTP/HTTPS URLs as valid (no network check performed).
                if (address.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                    address.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                // Resolve relative paths against the document's directory.
                string resolvedPath = address;
                if (!Path.IsPathRooted(address))
                {
                    string docDirectory = Path.GetDirectoryName(docPath) ?? string.Empty;
                    resolvedPath = Path.Combine(docDirectory, address);
                }

                // If the file does not exist, record it as a broken link.
                if (!File.Exists(resolvedPath))
                {
                    brokenLinks.Add(address);
                }
            }

            // -----------------------------------------------------------------
            // Step 3: Generate a simple report of the broken hyperlinks.
            // -----------------------------------------------------------------
            using (StreamWriter writer = new StreamWriter(reportPath, false))
            {
                if (brokenLinks.Count == 0)
                {
                    writer.WriteLine("No broken hyperlinks were found.");
                    Console.WriteLine("No broken hyperlinks were found.");
                }
                else
                {
                    writer.WriteLine("Broken Hyperlinks Report");
                    writer.WriteLine("========================");
                    foreach (string link in brokenLinks)
                    {
                        writer.WriteLine($"- {link}");
                        Console.WriteLine($"Broken link detected: {link}");
                    }
                }
            }

            // Clean up the temporary existing file created for the example.
            if (File.Exists(existingFileName))
                File.Delete(existingFileName);
        }
    }
}
