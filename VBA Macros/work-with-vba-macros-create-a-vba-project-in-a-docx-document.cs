using System;
using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a new VBA project and give it a name.
        VbaProject project = new VbaProject();
        project.Name = "MyAsposeProject";
        doc.VbaProject = project;

        // Create a procedural module that contains a simple macro.
        VbaModule module = new VbaModule();
        module.Name = "MyModule";
        module.Type = VbaModuleType.ProceduralModule;
        module.SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from VBA!""
End Sub";

        // Add the module to the VBA project.
        doc.VbaProject.Modules.Add(module);

        // Save the document as a macro‑enabled file (.docm).
        doc.Save("VbaProjectExample.docm");
    }
}
