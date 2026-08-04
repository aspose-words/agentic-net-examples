using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Define paths for the temporary VBA module file and the output document.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        string basFilePath = Path.Combine(artifactsDir, "SampleModule.bas");
        string outputDocPath = Path.Combine(artifactsDir, "DocumentWithImportedModule.docm");

        // Create a simple VBA module source and write it to a .bas file.
        string vbaSource = @"
Sub HelloWorld()
    MsgBox ""Hello from imported VBA module!""
End Sub
";
        File.WriteAllText(basFilePath, vbaSource);

        // Create a new blank Word document.
        Document doc = new Document();

        // Ensure the document has a VBA project.
        VbaProject vbaProject = new VbaProject();
        vbaProject.Name = "ImportedProject";
        doc.VbaProject = vbaProject;

        // Load the VBA source code from the .bas file.
        string importedSource = File.ReadAllText(basFilePath);

        // Create a new VBA module, set its name, type, and source code.
        VbaModule vbaModule = new VbaModule();
        vbaModule.Name = "ImportedModule";
        vbaModule.Type = VbaModuleType.ProceduralModule;
        vbaModule.SourceCode = importedSource;

        // Add the module to the VBA project.
        doc.VbaProject.Modules.Add(vbaModule);

        // Save the document in a macro‑enabled format.
        doc.Save(outputDocPath);
    }
}
