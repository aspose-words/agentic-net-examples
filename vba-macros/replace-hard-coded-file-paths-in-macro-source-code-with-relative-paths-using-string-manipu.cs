using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has a VBA project.
        VbaProject vbaProject = new VbaProject();
        vbaProject.Name = "SampleProject";
        doc.VbaProject = vbaProject;

        // Create a VBA module with hard‑coded absolute file paths.
        VbaModule module = new VbaModule();
        module.Name = "PathMacro";
        module.Type = VbaModuleType.ProceduralModule;
        module.SourceCode = @"
Sub InsertImage()
    Dim imgPath As String
    imgPath = ""C:\Data\Images\picture.png""
    ' Code that uses imgPath...
End Sub
";
        // Add the module to the project.
        doc.VbaProject.Modules.Add(module);

        // Replace absolute paths with relative paths in all modules.
        foreach (VbaModule mod in doc.VbaProject.Modules)
        {
            // Guard against null source code.
            string source = mod.SourceCode ?? string.Empty;

            // Example: replace any occurrence of "C:\Data\Images\" with "Images\".
            // This is a simple string manipulation; more complex logic can use Path methods.
            string oldPrefix = @"C:\Data\Images\";
            string newPrefix = @"Images\"; // Relative path.

            if (source.Contains(oldPrefix))
            {
                source = source.Replace(oldPrefix, newPrefix);
                mod.SourceCode = source;
            }
        }

        // Save the document as a macro‑enabled file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ModifiedMacro.docm");
        doc.Save(outputPath);
    }
}
