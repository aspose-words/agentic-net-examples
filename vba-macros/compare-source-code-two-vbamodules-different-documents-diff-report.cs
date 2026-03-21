using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

class VbaModuleDiff
{
    static void Main()
    {
        // Paths to the two documents containing VBA projects
        string docPath1 = @"C:\Docs\Document1.docm";
        string docPath2 = @"C:\Docs\Document2.docm";

        // Ensure the files exist; if not, create minimal .docm files so the example can run.
        docPath1 = EnsureDocumentExists(docPath1);
        docPath2 = EnsureDocumentExists(docPath2);

        // Load the documents
        Document doc1 = new Document(docPath1);
        Document doc2 = new Document(docPath2);

        // Access VBA projects (may be null for empty documents)
        VbaProject vbaProject1 = doc1.VbaProject;
        VbaProject vbaProject2 = doc2.VbaProject;

        if (vbaProject1 == null || vbaProject2 == null)
        {
            Console.WriteLine("One or both documents do not contain a VBA project. No comparison performed.");
            return;
        }

        // Build a lookup for modules in the second document by name
        var modules2ByName = new Dictionary<string, VbaModule>(StringComparer.OrdinalIgnoreCase);
        foreach (VbaModule mod in vbaProject2.Modules)
            modules2ByName[mod.Name] = mod;

        // Iterate through modules of the first document and compare source code
        foreach (VbaModule mod1 in vbaProject1.Modules)
        {
            Console.WriteLine($"--- Comparing module: {mod1.Name} ---");

            if (!modules2ByName.TryGetValue(mod1.Name, out VbaModule mod2))
            {
                Console.WriteLine("Module not found in second document.");
                continue;
            }

            string source1 = mod1.SourceCode ?? string.Empty;
            string source2 = mod2.SourceCode ?? string.Empty;

            if (source1 == source2)
            {
                Console.WriteLine("No differences found.");
                continue;
            }

            // Simple line‑by‑line diff
            string[] lines1 = source1.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            string[] lines2 = source2.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            int maxLines = Math.Max(lines1.Length, lines2.Length);

            for (int i = 0; i < maxLines; i++)
            {
                string line1 = i < lines1.Length ? lines1[i] : string.Empty;
                string line2 = i < lines2.Length ? lines2[i] : string.Empty;

                if (line1 != line2)
                {
                    Console.WriteLine($"Line {i + 1} differs:");
                    Console.WriteLine($"  Doc1: {line1}");
                    Console.WriteLine($"  Doc2: {line2}");
                }
            }
        }

        // Save a simple report to a writable location
        string reportPath = Path.Combine(Path.GetTempPath(), "VbaDiffReport.txt");
        using (StreamWriter writer = new StreamWriter(reportPath))
        {
            writer.WriteLine("VBA Module Diff Report");
            writer.WriteLine($"Generated on {DateTime.Now}");
            writer.WriteLine();
            writer.WriteLine("See console output for detailed differences.");
        }

        Console.WriteLine($"Diff report saved to: {reportPath}");
    }

    // Creates a minimal .docm file if the specified path does not exist and returns the path to use.
    private static string EnsureDocumentExists(string path)
    {
        if (File.Exists(path))
            return path;

        try
        {
            // Ensure directory exists
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);

            // Create a new blank document and save as .docm
            Document doc = new Document();
            doc.Save(path, SaveFormat.Docm);
            Console.WriteLine($"Created placeholder document at: {path}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to create placeholder document at '{path}': {ex.Message}");
            // Fallback to a temporary in‑memory document saved to a temp file
            string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docm");
            new Document().Save(tempPath, SaveFormat.Docm);
            Console.WriteLine($"Using temporary document at: {tempPath}");
            return tempPath;
        }

        return path;
    }
}
