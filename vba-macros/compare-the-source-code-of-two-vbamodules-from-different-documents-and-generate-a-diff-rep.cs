using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the two macro-enabled documents.
        string docPath1 = Path.Combine(outputDir, "Document1.docm");
        string docPath2 = Path.Combine(outputDir, "Document2.docm");

        // Create first document with a VBA module.
        Document doc1 = new Document();
        VbaProject project1 = new VbaProject();
        project1.Name = "Project1";
        doc1.VbaProject = project1;

        VbaModule module1 = new VbaModule();
        module1.Name = "SampleModule";
        module1.Type = VbaModuleType.ProceduralModule;
        module1.SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from Document 1!""
End Sub
";
        doc1.VbaProject.Modules.Add(module1);
        doc1.Save(docPath1);

        // Create second document with a slightly different VBA module.
        Document doc2 = new Document();
        VbaProject project2 = new VbaProject();
        project2.Name = "Project2";
        doc2.VbaProject = project2;

        VbaModule module2 = new VbaModule();
        module2.Name = "SampleModule";
        module2.Type = VbaModuleType.ProceduralModule;
        module2.SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from Document 2!""
    Debug.Print ""Additional line""
End Sub
";
        doc2.VbaProject.Modules.Add(module2);
        doc2.Save(docPath2);

        // Load the documents back (simulating external files).
        Document loadedDoc1 = new Document(docPath1);
        Document loadedDoc2 = new Document(docPath2);

        // Retrieve VBA projects.
        VbaProject vbaProject1 = loadedDoc1.VbaProject;
        VbaProject vbaProject2 = loadedDoc2.VbaProject;

        // Guard against missing VBA projects.
        if (vbaProject1 == null || vbaProject2 == null)
        {
            Console.WriteLine("One of the documents does not contain a VBA project.");
            return;
        }

        // Compare modules by name.
        foreach (VbaModule mod1 in vbaProject1.Modules)
        {
            // Find matching module in the second document.
            VbaModule mod2 = vbaProject2.Modules[mod1.Name];
            if (mod2 == null)
            {
                Console.WriteLine($"Module '{mod1.Name}' exists only in Document 1.");
                continue;
            }

            // Ensure source code strings are not null.
            string source1 = mod1.SourceCode ?? string.Empty;
            string source2 = mod2.SourceCode ?? string.Empty;

            // Split source code into lines for simple diff.
            string[] lines1 = source1.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            string[] lines2 = source2.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            int maxLines = Math.Max(lines1.Length, lines2.Length);
            bool differencesFound = false;

            Console.WriteLine($"Diff for module '{mod1.Name}':");
            for (int i = 0; i < maxLines; i++)
            {
                string line1 = i < lines1.Length ? lines1[i] : string.Empty;
                string line2 = i < lines2.Length ? lines2[i] : string.Empty;

                if (!string.Equals(line1.Trim(), line2.Trim(), StringComparison.Ordinal))
                {
                    differencesFound = true;
                    Console.WriteLine($"  Line {i + 1}:");
                    Console.WriteLine($"    Document1: {line1}");
                    Console.WriteLine($"    Document2: {line2}");
                }
            }

            if (!differencesFound)
                Console.WriteLine("  No differences found.");

            Console.WriteLine(); // Blank line between modules.
        }

        // Identify modules that exist only in Document 2.
        foreach (VbaModule mod2 in vbaProject2.Modules)
        {
            if (vbaProject1.Modules[mod2.Name] == null)
                Console.WriteLine($"Module '{mod2.Name}' exists only in Document 2.");
        }
    }
}
