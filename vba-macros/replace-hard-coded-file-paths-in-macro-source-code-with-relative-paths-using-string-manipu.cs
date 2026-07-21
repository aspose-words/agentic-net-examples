using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Define folder for output files (relative to the executable location).
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // Step 1: Create a new macro‑enabled document and add a VBA module.
        // -----------------------------------------------------------------
        Document doc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject vbaProject = new VbaProject();
        vbaProject.Name = "SampleProject";
        doc.VbaProject = vbaProject;

        // Create a VBA module with hard‑coded absolute file paths.
        VbaModule module = new VbaModule();
        module.Name = "PathMacro";
        module.Type = VbaModuleType.ProceduralModule;
        module.SourceCode = @"
Sub OpenFile()
    Dim filePath As String
    filePath = ""C:\Data\input.txt""
    ' Open the file (placeholder code)
    MsgBox ""Opening: "" & filePath
End Sub
";

        // Add the module to the VBA project.
        doc.VbaProject.Modules.Add(module);

        // Save the document as a macro‑enabled .docm file.
        string originalPath = Path.Combine(outputDir, "MacroDocument.docm");
        doc.Save(originalPath);

        // -----------------------------------------------------------------
        // Step 2: Load the saved document and replace absolute paths with
        //         relative paths in the VBA source code.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(originalPath);

        // Ensure the document actually contains macros.
        if (loadedDoc.HasMacros && loadedDoc.VbaProject != null)
        {
            foreach (VbaModule mod in loadedDoc.VbaProject.Modules)
            {
                // Guard against null source code.
                string source = mod.SourceCode ?? string.Empty;

                // Replace the hard‑coded absolute path with a relative one.
                // Example: "C:\Data\" -> ".\Data\"
                source = source.Replace(@"C:\Data\", @".\Data\");

                // Assign the modified source back to the module.
                mod.SourceCode = source;
            }

            // Save the updated document.
            string updatedPath = Path.Combine(outputDir, "MacroDocument_Updated.docm");
            loadedDoc.Save(updatedPath);
        }

        // Indicate completion (no interactive input required).
        Console.WriteLine("Macro path replacement completed.");
    }
}
