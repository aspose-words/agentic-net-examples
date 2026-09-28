using System;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has a VBA project.
        doc.VbaProject = new VbaProject();

        // Sample VBA macro source containing a hard‑coded absolute file path.
        string originalVba =
@"Sub Test()
    Dim path As String
    path = ""C:\Data\file.txt""
    MsgBox path
End Sub";

        // Create a new VBA module, set its name and source code, then add it to the project.
        VbaModule module = new VbaModule();
        module.Name = "Module1";
        module.SourceCode = originalVba;
        doc.VbaProject.Modules.Add(module);

        // Retrieve the module (guard against null source code).
        VbaModule targetModule = doc.VbaProject.Modules["Module1"];
        string sourceCode = targetModule?.SourceCode ?? string.Empty;

        // Replace the absolute path with a relative path (just the file name).
        string absolutePath = @"C:\Data\file.txt";
        string relativePath = "file.txt";
        string updatedSource = sourceCode.Replace(absolutePath, relativePath);

        // Update the module's source code.
        if (targetModule != null)
        {
            targetModule.SourceCode = updatedSource;
        }

        // Save the document in macro‑enabled format.
        string outputPath = "output.docm";
        doc.Save(outputPath);

        // Reload the document to verify the change.
        Document loadedDoc = new Document(outputPath);
        VbaModule loadedModule = loadedDoc.VbaProject?.Modules["Module1"];
        string loadedSource = loadedModule?.SourceCode ?? string.Empty;

        // Simple validation output.
        bool containsRelative = loadedSource.Contains(relativePath);
        bool containsAbsolute = loadedSource.Contains(absolutePath);
        Console.WriteLine($"Relative path present: {containsRelative}");
        Console.WriteLine($"Absolute path present: {containsAbsolute}");
    }
}
