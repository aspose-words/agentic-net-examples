using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = "Artifacts";
        Directory.CreateDirectory(artifactsDir);

        // Paths for the original and updated documents.
        string originalPath = Path.Combine(artifactsDir, "MacroWithAbsolutePath.docm");
        string updatedPath = Path.Combine(artifactsDir, "MacroWithRelativePath.docm");

        // -------------------------------------------------
        // Create a new macro‑enabled document with a VBA module that contains a hard‑coded absolute path.
        // -------------------------------------------------
        Document doc = new Document();

        VbaProject project = new VbaProject();
        project.Name = "SampleProject";
        doc.VbaProject = project;

        VbaModule module = new VbaModule();
        module.Name = "PathModule";
        module.Type = VbaModuleType.ProceduralModule;
        module.SourceCode = @"
Sub OpenFile()
    Dim filePath As String
    filePath = ""C:\Data\myfile.txt""
    MsgBox ""Opening "" & filePath
End Sub
";
        doc.VbaProject.Modules.Add(module);

        // Save the document containing the absolute path.
        doc.Save(originalPath);

        // -------------------------------------------------
        // Load the document and replace the absolute path with a relative one.
        // -------------------------------------------------
        Document loadedDoc = new Document(originalPath);

        if (loadedDoc.HasMacros && loadedDoc.VbaProject != null)
        {
            foreach (VbaModule mod in loadedDoc.VbaProject.Modules)
            {
                // Guard against null source code.
                string source = mod.SourceCode ?? string.Empty;

                // Replace the hard‑coded part of the path.
                string updatedSource = source.Replace(@"C:\Data\", @".\Data\");

                mod.SourceCode = updatedSource;
            }
        }

        // Save the modified document.
        loadedDoc.Save(updatedPath);

        // Simple console output to indicate completion.
        Console.WriteLine("Original macro saved to: " + originalPath);
        Console.WriteLine("Updated macro saved to: " + updatedPath);
    }
}
