using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Define file names for the two macro-enabled documents.
        string docPath1 = "Doc1.docm";
        string docPath2 = "Doc2.docm";

        // Create the first document with a VBA module.
        Document doc1 = new Document();
        VbaProject project1 = new VbaProject();
        project1.Name = "Project1";
        doc1.VbaProject = project1;

        VbaModule module1 = new VbaModule();
        module1.Name = "Module1";
        module1.Type = VbaModuleType.ProceduralModule;
        module1.SourceCode = @"Sub Test()
    MsgBox ""Hello""
End Sub";
        doc1.VbaProject.Modules.Add(module1);
        doc1.Save(docPath1);

        // Create the second document with a slightly different VBA module.
        Document doc2 = new Document();
        VbaProject project2 = new VbaProject();
        project2.Name = "Project2";
        doc2.VbaProject = project2;

        VbaModule module2 = new VbaModule();
        module2.Name = "Module1";
        module2.Type = VbaModuleType.ProceduralModule;
        module2.SourceCode = @"Sub Test()
    MsgBox ""Hello World""
End Sub";
        doc2.VbaProject.Modules.Add(module2);
        doc2.Save(docPath2);

        // Load the two documents back from disk.
        Document loadedDoc1 = new Document(docPath1);
        Document loadedDoc2 = new Document(docPath2);

        // Retrieve the VBA modules (by name) from each document.
        VbaModule loadedModule1 = loadedDoc1.VbaProject?.Modules["Module1"];
        VbaModule loadedModule2 = loadedDoc2.VbaProject?.Modules["Module1"];

        // Guard against missing modules or null source code.
        string source1 = loadedModule1?.SourceCode ?? string.Empty;
        string source2 = loadedModule2?.SourceCode ?? string.Empty;

        // Split source code into lines for a simple line‑by‑line diff.
        string[] lines1 = source1.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
        string[] lines2 = source2.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
        int maxLines = Math.Max(lines1.Length, lines2.Length);

        List<string> diffReport = new List<string>();
        diffReport.Add("=== VBA Module Diff Report ===");
        diffReport.Add($"Comparing module \"{loadedModule1?.Name ?? "N/A"}\" between documents.");

        for (int i = 0; i < maxLines; i++)
        {
            string line1 = i < lines1.Length ? lines1[i] : null;
            string line2 = i < lines2.Length ? lines2[i] : null;

            if (line1 == line2)
                continue; // No difference on this line.

            diffReport.Add($"Line {i + 1}:");
            diffReport.Add($"  Doc1: {(line1 ?? "<no line>")}");
            diffReport.Add($"  Doc2: {(line2 ?? "<no line>")}");
        }

        // Output the diff report to the console.
        foreach (string line in diffReport)
        {
            Console.WriteLine(line);
        }

        // Clean up generated files (optional).
        // File.Delete(docPath1);
        // File.Delete(docPath2);
    }
}
