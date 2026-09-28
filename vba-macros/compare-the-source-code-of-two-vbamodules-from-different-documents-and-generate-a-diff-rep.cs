using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Create first macro-enabled document with a VBA module.
        Document doc1 = new Document();
        doc1.VbaProject = new VbaProject();

        VbaModule module1 = new VbaModule
        {
            Name = "SampleModule",
            SourceCode = @"Sub Hello()
    MsgBox ""Hello from Document 1""
End Sub"
        };
        doc1.VbaProject.Modules.Add(module1);

        string file1 = "doc1.docm";
        doc1.Save(file1, SaveFormat.Docm);

        // Create second macro-enabled document with a slightly different VBA module.
        Document doc2 = new Document();
        doc2.VbaProject = new VbaProject();

        VbaModule module2 = new VbaModule
        {
            Name = "SampleModule",
            SourceCode = @"Sub Hello()
    MsgBox ""Hello from Document 2""
    Debug.Print ""Additional line""
End Sub"
        };
        doc2.VbaProject.Modules.Add(module2);

        string file2 = "doc2.docm";
        doc2.Save(file2, SaveFormat.Docm);

        // Reload documents to ensure they are read from disk.
        Document loadedDoc1 = new Document(file1);
        Document loadedDoc2 = new Document(file2);

        // Retrieve the VBA modules (guard against missing modules).
        string source1 = string.Empty;
        string source2 = string.Empty;

        if (loadedDoc1.VbaProject?.Modules?["SampleModule"] != null)
        {
            source1 = loadedDoc1.VbaProject.Modules["SampleModule"].SourceCode ?? string.Empty;
        }

        if (loadedDoc2.VbaProject?.Modules?["SampleModule"] != null)
        {
            source2 = loadedDoc2.VbaProject.Modules["SampleModule"].SourceCode ?? string.Empty;
        }

        // Split source code into lines for comparison.
        string[] lines1 = source1.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
        string[] lines2 = source2.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

        // Build simple diff report.
        var onlyInFirst = new List<string>();
        var onlyInSecond = new List<string>();
        var set1 = new HashSet<string>(lines1);
        var set2 = new HashSet<string>(lines2);

        foreach (string line in lines2)
        {
            if (!set1.Contains(line))
                onlyInSecond.Add(line);
        }

        foreach (string line in lines1)
        {
            if (!set2.Contains(line))
                onlyInFirst.Add(line);
        }

        // Output diff report.
        Console.WriteLine("=== Diff Report for VBA Module 'SampleModule' ===");
        Console.WriteLine();

        Console.WriteLine($"Lines only in first document ({file1}):");
        if (onlyInFirst.Count == 0)
        {
            Console.WriteLine("  (none)");
        }
        else
        {
            foreach (var line in onlyInFirst)
                Console.WriteLine("- " + line);
        }

        Console.WriteLine();

        Console.WriteLine($"Lines only in second document ({file2}):");
        if (onlyInSecond.Count == 0)
        {
            Console.WriteLine("  (none)");
        }
        else
        {
            foreach (var line in onlyInSecond)
                Console.WriteLine("+ " + line);
        }

        // Clean up temporary files.
        TryDelete(file1);
        TryDelete(file2);
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
                File.Delete(path);
        }
        catch
        {
            // Ignored – cleanup failure should not affect program outcome.
        }
    }
}
