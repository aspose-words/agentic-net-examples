using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class ImportVbaModuleExample
{
    public static void Main()
    {
        // Path for the temporary VBA module file (.bas)
        string basFilePath = "SampleModule.bas";

        // Create a simple VBA module source code and write it to the .bas file
        string vbaSource = @"Attribute VB_Name = ""SampleModule""
Sub HelloWorld()
    MsgBox ""Hello from VBA!""
End Sub";
        File.WriteAllText(basFilePath, vbaSource);

        // Create a new blank document
        Document doc = new Document();

        // Ensure the document has a VBA project; create one if it does not exist
        if (doc.VbaProject == null)
        {
            VbaProject project = new VbaProject();
            project.Name = "ImportedProject";
            doc.VbaProject = project;
        }

        // Read the VBA source code from the external .bas file
        string importedSource = File.ReadAllText(basFilePath);

        // Create a new VBA module, set its name, type, and source code
        VbaModule module = new VbaModule();
        module.Name = "MyImportedModule";               // Desired module name
        module.Type = VbaModuleType.ProceduralModule;   // Typical procedural module
        module.SourceCode = importedSource;              // Assign the imported source code

        // Add the module to the document's VBA project
        doc.VbaProject.Modules.Add(module);

        // Save the document in a macro‑enabled format
        string outputPath = "ImportedMacro.docm";
        doc.Save(outputPath);

        // Clean up the temporary .bas file (optional)
        if (File.Exists(basFilePath))
        {
            File.Delete(basFilePath);
        }

        // Indicate completion (no interactive prompts)
        Console.WriteLine("VBA module imported and document saved to: " + outputPath);
    }
}
