using System;
using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Create a blank Word document.
        Document doc = new Document();

        // Initialize a new VBA project and set its name.
        VbaProject project = new VbaProject();
        project.Name = "MyAsposeProject";
        doc.VbaProject = project;

        // Create a procedural VBA module with sample macro code.
        VbaModule module = new VbaModule();
        module.Name = "MyModule";
        module.Type = VbaModuleType.ProceduralModule;
        module.SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from VBA!""
End Sub";

        // Add the module to the VBA project.
        doc.VbaProject.Modules.Add(module);

        // Save the document as a macro‑enabled file.
        doc.Save("VbaProjectInDocm.docm");
    }
}
