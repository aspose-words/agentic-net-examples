using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Directory for temporary files
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the two macro-enabled documents
        string docPath1 = Path.Combine(outputDir, "Doc1.docm");
        string docPath2 = Path.Combine(outputDir, "Doc2.docm");

        // Create first document with a VBA module
        Document doc1 = new Document();
        EnsureVbaProject(doc1);
        VbaModule module1 = new VbaModule
        {
            Name = "SampleModule",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from Document 1!""
End Sub

Sub AddNumbers()
    Dim a As Integer
    Dim b As Integer
    a = 5
    b = 10
    MsgBox a + b
End Sub"
        };
        doc1.VbaProject.Modules.Add(module1);
        doc1.Save(docPath1);

        // Create second document with a slightly different VBA module
        Document doc2 = new Document();
        EnsureVbaProject(doc2);
        VbaModule module2 = new VbaModule
        {
            Name = "SampleModule",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from Document 2!""
End Sub

Sub AddNumbers()
    Dim a As Integer
    Dim b As Integer
    a = 7
    b = 3
    MsgBox a + b
End Sub

Sub NewProcedure()
    MsgBox ""New procedure in Document 2""
End Sub"
        };
        doc2.VbaProject.Modules.Add(module2);
        doc2.Save(docPath2);

        // Load the documents (simulating separate sources)
        Document loadedDoc1 = new Document(docPath1);
        Document loadedDoc2 = new Document(docPath2);

        // Retrieve the first module from each document (by name)
        VbaModule vbaModule1 = loadedDoc1.VbaProject?.Modules["SampleModule"];
        VbaModule vbaModule2 = loadedDoc2.VbaProject?.Modules["SampleModule"];

        // Guard against null source code
        string source1 = vbaModule1?.SourceCode ?? string.Empty;
        string source2 = vbaModule2?.SourceCode ?? string.Empty;

        // Generate a simple line‑by‑line diff report
        List<string> diffReport = GenerateDiffReport(source1, source2);

        // Output the diff report
        Console.WriteLine("Diff report between VBA modules:");
        foreach (string line in diffReport)
        {
            Console.WriteLine(line);
        }
    }

    // Ensures the document has a VbaProject; creates one if missing
    private static void EnsureVbaProject(Document doc)
    {
        if (doc.VbaProject == null)
        {
            VbaProject project = new VbaProject
            {
                Name = "AsposeProject"
            };
            doc.VbaProject = project;
        }
    }

    // Simple diff: lines prefixed with '+' (added), '-' (removed), or ' ' (unchanged)
    private static List<string> GenerateDiffReport(string text1, string text2)
    {
        string[] lines1 = text1.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
        string[] lines2 = text2.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

        var report = new List<string>();
        int i = 0, j = 0;
        while (i < lines1.Length || j < lines2.Length)
        {
            if (i < lines1.Length && j < lines2.Length)
            {
                if (lines1[i] == lines2[j])
                {
                    report.Add("  " + lines1[i]);
                    i++; j++;
                }
                else
                {
                    report.Add("- " + lines1[i]);
                    report.Add("+ " + lines2[j]);
                    i++; j++;
                }
            }
            else if (i < lines1.Length)
            {
                report.Add("- " + lines1[i]);
                i++;
            }
            else // j < lines2.Length
            {
                report.Add("+ " + lines2[j]);
                j++;
            }
        }
        return report;
    }
}
