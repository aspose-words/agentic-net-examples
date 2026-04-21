using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Define output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Path for the initial macro-enabled document with hard‑coded paths.
        string absoluteDocPath = Path.Combine(outputDir, "MacroWithAbsolutePath.docm");
        // Path for the document after converting to relative paths.
        string relativeDocPath = Path.Combine(outputDir, "MacroWithRelativePath.docm");

        // -----------------------------------------------------------------
        // 1. Create a new blank document and a VBA project.
        // -----------------------------------------------------------------
        Document doc = new Document();

        VbaProject vbaProject = new VbaProject
        {
            Name = "SampleProject"
        };
        doc.VbaProject = vbaProject;

        // -----------------------------------------------------------------
        // 2. Add a VBA module containing a hard‑coded absolute file path.
        // -----------------------------------------------------------------
        VbaModule module = new VbaModule
        {
            Name = "PathMacro",
            Type = VbaModuleType.ProceduralModule,
            // Example VBA code with an absolute Windows path.
            SourceCode = @"
Sub OpenFile()
    Dim filePath As String
    filePath = ""C:\Data\MyFile.txt""
    MsgBox ""Opening: "" & filePath
End Sub"
        };

        doc.VbaProject.Modules.Add(module);

        // Save the document with the absolute path macro.
        doc.Save(absoluteDocPath);

        // -----------------------------------------------------------------
        // 3. Load the saved document and replace absolute paths with relative ones.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(absoluteDocPath);

        // Ensure the document actually contains macros.
        if (loadedDoc.HasMacros && loadedDoc.VbaProject != null)
        {
            foreach (VbaModule mod in loadedDoc.VbaProject.Modules)
            {
                // Guard against null source code.
                string source = mod.SourceCode ?? string.Empty;

                // Replace the hard‑coded absolute part with a relative path.
                // Example: replace "C:\Data\" with ".\Data\"
                string updatedSource = source.Replace(@"C:\Data\", @".\Data\");

                // Apply the modified source back to the module.
                mod.SourceCode = updatedSource;
            }

            // Save the modified document.
            loadedDoc.Save(relativeDocPath);
        }

        // -----------------------------------------------------------------
        // 4. Simple verification output (optional).
        // -----------------------------------------------------------------
        Console.WriteLine("Macro with absolute path saved to:");
        Console.WriteLine(absoluteDocPath);
        Console.WriteLine("Macro with relative path saved to:");
        Console.WriteLine(relativeDocPath);
    }
}
