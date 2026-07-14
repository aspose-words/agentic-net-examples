using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Define output directory.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Step 1: Create a new macro‑enabled document and add a VBA module with an absolute file path.
        Document doc = new Document();

        // Create a new VBA project if none exists.
        VbaProject vbaProject = new VbaProject();
        vbaProject.Name = "SampleProject";
        doc.VbaProject = vbaProject;

        // Create a procedural VBA module.
        VbaModule module = new VbaModule();
        module.Name = "PathMacro";
        module.Type = VbaModuleType.ProceduralModule;

        // Sample VBA code containing a hard‑coded absolute path.
        // The macro simply opens a file using the absolute path.
        module.SourceCode = @"
Sub OpenFile()
    Dim filePath As String
    filePath = ""C:\Data\Files\myfile.txt""
    MsgBox ""Opening: "" & filePath
End Sub
";

        // Add the module to the project.
        doc.VbaProject.Modules.Add(module);

        // Save the document with the absolute path in the macro.
        string absoluteMacroPath = Path.Combine(artifactsDir, "MacroWithAbsolutePaths.docm");
        doc.Save(absoluteMacroPath);

        // Step 2: Load the saved document and replace the absolute path with a relative one.
        Document loadedDoc = new Document(absoluteMacroPath);

        // Ensure the document actually has a VBA project.
        if (loadedDoc.HasMacros && loadedDoc.VbaProject != null)
        {
            foreach (VbaModule vbaModule in loadedDoc.VbaProject.Modules)
            {
                // Guard against null source code.
                string source = vbaModule.SourceCode ?? string.Empty;

                // Replace the hard‑coded absolute directory with a relative path.
                // Example: "C:\Data\Files\" -> "./Files/"
                string absoluteDir = @"C:\Data\Files\";
                string relativeDir = "./Files/";

                if (source.Contains(absoluteDir))
                {
                    source = source.Replace(absoluteDir, relativeDir);
                    vbaModule.SourceCode = source;
                }
            }
        }

        // Save the modified document.
        string relativeMacroPath = Path.Combine(artifactsDir, "MacroWithRelativePaths.docm");
        loadedDoc.Save(relativeMacroPath);

        // Output the paths of the generated files (no user interaction required).
        Console.WriteLine("Created document with absolute paths: " + absoluteMacroPath);
        Console.WriteLine("Created document with relative paths: " + relativeMacroPath);
    }
}
