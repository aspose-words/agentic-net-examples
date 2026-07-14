using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Directory for temporary files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create first macro-enabled document with a VBA module.
        Document doc1 = new Document();
        VbaProject project1 = new VbaProject { Name = "Project1" };
        doc1.VbaProject = project1;

        VbaModule module1 = new VbaModule
        {
            Name = "ModuleA",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"Sub Hello()
    MsgBox ""Hello from Doc1!""
End Sub"
        };
        doc1.VbaProject.Modules.Add(module1);
        string doc1Path = Path.Combine(artifactsDir, "Doc1.docm");
        doc1.Save(doc1Path);

        // Create second macro-enabled document with a slightly different VBA module.
        Document doc2 = new Document();
        VbaProject project2 = new VbaProject { Name = "Project2" };
        doc2.VbaProject = project2;

        VbaModule module2 = new VbaModule
        {
            Name = "ModuleA",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"Sub Hello()
    MsgBox ""Hello from Doc2!""
    ' Added comment
End Sub"
        };
        doc2.VbaProject.Modules.Add(module2);
        string doc2Path = Path.Combine(artifactsDir, "Doc2.docm");
        doc2.Save(doc2Path);

        // Load the two documents for comparison.
        Document loadedDoc1 = new Document(doc1Path);
        Document loadedDoc2 = new Document(doc2Path);

        // Retrieve the modules (by name) – if not found, treat as empty source.
        string moduleName = "ModuleA";
        string source1 = GetModuleSource(loadedDoc1, moduleName);
        string source2 = GetModuleSource(loadedDoc2, moduleName);

        // Generate a simple line‑by‑line diff report.
        string[] lines1 = source1.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
        string[] lines2 = source2.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
        int maxLines = Math.Max(lines1.Length, lines2.Length);

        Console.WriteLine("=== VBA Module Diff Report ===");
        for (int i = 0; i < maxLines; i++)
        {
            string line1 = i < lines1.Length ? lines1[i] : string.Empty;
            string line2 = i < lines2.Length ? lines2[i] : string.Empty;

            if (line1 != line2)
            {
                Console.WriteLine($"Line {i + 1}:");
                Console.WriteLine($"  Doc1: {line1}");
                Console.WriteLine($"  Doc2: {line2}");
            }
        }

        // Clean up temporary files (optional).
        // File.Delete(doc1Path);
        // File.Delete(doc2Path);
    }

    // Helper to safely obtain a module's source code.
    private static string GetModuleSource(Document doc, string moduleName)
    {
        if (doc.HasMacros && doc.VbaProject != null)
        {
            VbaModule module = doc.VbaProject.Modules[moduleName];
            if (module != null && !string.IsNullOrEmpty(module.SourceCode))
                return module.SourceCode;
        }
        return string.Empty;
    }
}
