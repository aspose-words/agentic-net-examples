using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Vba;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create first macro‑enabled document with a VBA module
        var doc1 = new Document();
        if (doc1.VbaProject == null)
        {
            doc1.VbaProject = new VbaProject();
        }

        var module1 = new VbaModule();
        module1.Name = "TestModule";
        module1.SourceCode = @"Sub Test()
    MsgBox ""Doc1 version""
End Sub";
        doc1.VbaProject.Modules.Add(module1);
        doc1.Save("doc1.docm", SaveFormat.Docm);

        // Create second macro‑enabled document with a slightly different VBA module
        var doc2 = new Document();
        if (doc2.VbaProject == null)
        {
            doc2.VbaProject = new VbaProject();
        }

        var module2 = new VbaModule();
        module2.Name = "TestModule";
        module2.SourceCode = @"Sub Test()
    MsgBox ""Doc2 version""
    Debug.Print ""Additional line""
End Sub";
        doc2.VbaProject.Modules.Add(module2);
        doc2.Save("doc2.docm", SaveFormat.Docm);

        // Retrieve the modules (by name) from each document
        var retrievedModule1 = doc1.VbaProject.Modules["TestModule"];
        var retrievedModule2 = doc2.VbaProject.Modules["TestModule"];

        // Guard against null source code
        string source1 = retrievedModule1?.SourceCode ?? string.Empty;
        string source2 = retrievedModule2?.SourceCode ?? string.Empty;

        // Generate and output the diff report
        string diffReport = GenerateDiffReport(source1, source2);
        Console.WriteLine(diffReport);
    }

    private static string GenerateDiffReport(string text1, string text2)
    {
        var lines1 = text1.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        var lines2 = text2.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        int maxLines = Math.Max(lines1.Length, lines2.Length);
        var sb = new StringBuilder();
        sb.AppendLine("Diff Report:");
        bool anyDifference = false;

        for (int i = 0; i < maxLines; i++)
        {
            string line1 = i < lines1.Length ? lines1[i] : string.Empty;
            string line2 = i < lines2.Length ? lines2[i] : string.Empty;

            if (!string.Equals(line1, line2, StringComparison.Ordinal))
            {
                anyDifference = true;
                sb.AppendLine($"Line {i + 1}:");
                sb.AppendLine($"  Doc1: {line1}");
                sb.AppendLine($"  Doc2: {line2}");
            }
        }

        if (!anyDifference)
        {
            sb.AppendLine("  No differences found.");
        }

        return sb.ToString();
    }
}
