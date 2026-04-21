using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Path for the temporary VBA module file.
        const string basFilePath = "SampleModule.bas";

        // Create a simple VBA macro and write it to the .bas file.
        const string vbaCode = 
@"Sub HelloWorld()
    MsgBox ""Hello from VBA!""
End Sub";

        File.WriteAllText(basFilePath, vbaCode);

        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has a VBA project.
        if (doc.VbaProject == null)
        {
            VbaProject project = new VbaProject();
            project.Name = "ImportedProject";
            doc.VbaProject = project;
        }

        // Read the VBA source code from the .bas file.
        string sourceCode = File.ReadAllText(basFilePath) ?? string.Empty;

        // Create a new VBA module and set its properties.
        VbaModule module = new VbaModule();
        module.Name = "ImportedModule";
        module.Type = VbaModuleType.ProceduralModule;
        module.SourceCode = sourceCode;

        // Add the module to the document's VBA project.
        doc.VbaProject.Modules.Add(module);

        // Save the document in a macro‑enabled format.
        const string outputPath = "DocumentWithImportedModule.docm";
        doc.Save(outputPath);

        // Simple verification output.
        Console.WriteLine($"Document saved to '{outputPath}'.");
        Console.WriteLine($"Module '{module.Name}' imported with source length {module.SourceCode.Length} characters.");
    }
}
