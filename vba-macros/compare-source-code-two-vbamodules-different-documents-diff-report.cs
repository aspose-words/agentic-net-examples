using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

class VbaModuleDiff
{
    // Entry point
    static void Main()
    {
        // Paths to the two documents containing VBA projects
        string docPath1 = @"C:\Docs\Document1.docm";
        string docPath2 = @"C:\Docs\Document2.docm";

        // Load the documents (lifecycle rule: use Document constructor)
        Document doc1 = new Document(docPath1);
        Document doc2 = new Document(docPath2);

        // Access VBA projects
        VbaProject vbaProject1 = doc1.VbaProject;
        VbaProject vbaProject2 = doc2.VbaProject;

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

        // Optionally, save a report file (using the allowed Save method)
        string reportPath = @"C:\Docs\VbaDiffReport.txt";
        using (StreamWriter writer = new StreamWriter(reportPath))
        {
            writer.WriteLine("VBA Module Diff Report");
            writer.WriteLine($"Generated on {DateTime.Now}");
            // Additional detailed report could be written here following the same logic as above.
        }

        Console.WriteLine($"Diff report saved to: {reportPath}");
    }
}
