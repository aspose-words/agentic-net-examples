using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Create a blank Word document.
        Document doc = new Document();

        // Create a new VBA project and give it a name.
        VbaProject vbaProject = new VbaProject();
        vbaProject.Name = "MyVbaProject";

        // Attach the VBA project to the document.
        doc.VbaProject = vbaProject;

        // Create a new procedural module.
        VbaModule module = new VbaModule();
        module.Name = "MyModule";
        module.Type = VbaModuleType.ProceduralModule;

        // Assign custom macro code to the module.
        module.SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from Aspose.Words!""
End Sub
";

        // Add the module to the VBA project.
        doc.VbaProject.Modules.Add(module);

        // Save the document in a macro‑enabled format.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CreatedMacro.docm");
        doc.Save(outputPath);

        // Output simple verification information.
        Console.WriteLine($"Document saved to: {outputPath}");
        Console.WriteLine($"VBA project contains {doc.VbaProject.Modules.Count} module(s).");
    }
}
