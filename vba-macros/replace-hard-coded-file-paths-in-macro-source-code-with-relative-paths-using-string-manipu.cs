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

        // VBA code with a hard‑coded absolute file path.
        string vbaCode = @"
Sub Example()
    Dim filePath As String
    filePath = ""C:\Data\MyFile.txt""
    MsgBox filePath
End Sub
";

        // Create a VBA module, set its name and source code, then add it to the project.
        VbaModule newModule = new VbaModule();
        newModule.Name = "Module1";
        newModule.SourceCode = vbaCode;
        doc.VbaProject.Modules.Add(newModule);

        // Retrieve the module.
        VbaModule module = doc.VbaProject.Modules["Module1"];
        string source = module.SourceCode ?? string.Empty;

        // Replace the absolute path with a relative path.
        string absolutePath = @"C:\Data\MyFile.txt";
        string relativePath = @"..\Data\MyFile.txt";
        string updatedSource = source.Replace(absolutePath, relativePath);

        // Update the module source code.
        module.SourceCode = updatedSource;

        // Save the document in macro‑enabled format.
        string fileName = "MacroDocument.docm";
        doc.Save(fileName);

        // Reload the document to verify the change.
        Document loadedDoc = new Document(fileName);
        VbaModule loadedModule = loadedDoc.VbaProject.Modules["Module1"];
        string finalSource = loadedModule.SourceCode ?? string.Empty;

        bool containsRelative = finalSource.Contains(relativePath);
        Console.WriteLine("Relative path replacement successful: " + containsRelative);
    }
}
