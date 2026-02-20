using System;
using Aspose.Words;
using Aspose.Words.Vba;

class VbaProjectExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject project = new VbaProject();
        project.Name = "Aspose.Project";
        doc.VbaProject = project;

        // Create a new procedural module with some macro source code.
        VbaModule module = new VbaModule();
        module.Name = "Aspose.Module";
        module.Type = VbaModuleType.ProceduralModule;
        module.SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from Aspose.Words VBA!""
End Sub
";

        // Add the module to the VBA project.
        doc.VbaProject.Modules.Add(module);

        // Save the document as a macro-enabled DOCM file.
        doc.Save("VbaProject.CreateVBAMacros.docm");
    }
}
