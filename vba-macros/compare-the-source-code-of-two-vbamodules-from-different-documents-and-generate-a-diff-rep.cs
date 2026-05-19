using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create first macro-enabled document with a VBA module.
        Document doc1 = new Document();
        VbaProject project1 = new VbaProject { Name = "Project1" };
        doc1.VbaProject = project1;

        VbaModule module1 = new VbaModule
        {
            Name = "Module1",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"Sub Test()
    MsgBox ""Hello""
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
            Name = "Module1",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"Sub Test()
    MsgBox ""Hello World""
End Sub"
        };
        doc2.VbaProject.Modules.Add(module2);
        string doc2Path = Path.Combine(artifactsDir, "Doc2.docm");
        doc2.Save(doc2Path);

        // Reload documents to simulate independent sources.
        Document loadedDoc1 = new Document(doc1Path);
        Document loadedDoc2 = new Document(doc2Path);

        // Retrieve modules (by name) from each document.
        VbaModule loadedModule1 = loadedDoc1.VbaProject?.Modules["Module1"];
        VbaModule loadedModule2 = loadedDoc2.VbaProject?.Modules["Module1"];

        // Guard against missing modules or null source code.
        string source1 = loadedModule1?.SourceCode ?? string.Empty;
        string source2 = loadedModule2?.SourceCode ?? string.Empty;

        // Split source code into lines for comparison.
        string[] lines1 = source1.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
        string[] lines2 = source2.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
        int maxLines = Math.Max(lines1.Length, lines2.Length);

        // Build diff report.
        List<string> diffReport = new List<string>();
        diffReport.Add("Diff report for VBA module 'Module1':");
        for (int i = 0; i < maxLines; i++)
        {
            string line1 = i < lines1.Length ? lines1[i] : string.Empty;
            string line2 = i < lines2.Length ? lines2[i] : string.Empty;

            if (!string.Equals(line1, line2, StringComparison.Ordinal))
            {
                diffReport.Add($"Line {i + 1}:");
                diffReport.Add($"  Doc1: \"{line1}\"");
                diffReport.Add($"  Doc2: \"{line2}\"");
            }
        }

        // If no differences were found, note that.
        if (diffReport.Count == 1)
            diffReport.Add("No differences found.");

        // Output the diff report to console.
        foreach (string line in diffReport)
            Console.WriteLine(line);
    }
}
