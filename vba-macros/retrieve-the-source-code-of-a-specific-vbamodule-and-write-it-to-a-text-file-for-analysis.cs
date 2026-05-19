using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class RetrieveVbaModuleSource
{
    public static void Main()
    {
        // Define file paths.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);
        string docPath = Path.Combine(outputDir, "SampleDocument.docm");
        string sourcePath = Path.Combine(outputDir, "ModuleSource.txt");

        // Create a new blank document.
        Document doc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject vbaProject = new VbaProject();
        vbaProject.Name = "SampleProject";
        doc.VbaProject = vbaProject;

        // Create a new VBA module with sample code.
        VbaModule vbaModule = new VbaModule();
        vbaModule.Name = "SampleModule";
        vbaModule.Type = VbaModuleType.ProceduralModule;
        vbaModule.SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from VBA!""
End Sub
";

        // Add the module to the VBA project.
        doc.VbaProject.Modules.Add(vbaModule);

        // Save the document in a macro‑enabled format.
        doc.Save(docPath, SaveFormat.Docm);

        // Load the document (optional, demonstrates loading from file).
        Document loadedDoc = new Document(docPath);

        // Access the VBA project and retrieve the specific module by name.
        VbaProject loadedProject = loadedDoc.VbaProject;
        VbaModule targetModule = loadedProject?.Modules["SampleModule"];

        // Guard against null source code.
        string sourceCode = targetModule?.SourceCode ?? string.Empty;

        // Write the source code to a text file.
        File.WriteAllText(sourcePath, sourceCode);
    }
}
