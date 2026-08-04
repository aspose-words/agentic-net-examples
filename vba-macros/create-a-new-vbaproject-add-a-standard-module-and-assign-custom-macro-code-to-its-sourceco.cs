using System;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject project = new VbaProject();
        project.Name = "Aspose.Project";
        doc.VbaProject = project;

        // Create a new standard (procedural) VBA module.
        VbaModule module = new VbaModule();
        module.Name = "MyStandardModule";
        module.Type = VbaModuleType.ProceduralModule;

        // Assign custom macro code to the module's SourceCode property.
        module.SourceCode = @"
Sub MyMacro()
    MsgBox ""Hello from VBA!""
End Sub
";

        // Add the module to the VBA project.
        doc.VbaProject.Modules.Add(module);

        // Save the document in a macro‑enabled format (.docm).
        doc.Save("VbaProject.CreateVBAMacros.docm");
    }
}
