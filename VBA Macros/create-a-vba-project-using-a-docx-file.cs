using System;
using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Create a blank Word document.
        Document doc = new Document();

        // Create a new VBA project and give it a name.
        VbaProject project = new VbaProject();
        project.Name = "MyAsposeProject";
        doc.VbaProject = project;

        // Create a procedural VBA module and add some macro source code.
        VbaModule module = new VbaModule();
        module.Name = "MyModule";
        module.Type = VbaModuleType.ProceduralModule;
        module.SourceCode = 
            "Sub HelloWorld()\n" +
            "    MsgBox \"Hello from VBA!\"\n" +
            "End Sub";

        // Add the module to the VBA project's module collection.
        doc.VbaProject.Modules.Add(module);

        // Save the document as a macro‑enabled .docm file.
        doc.Save("VbaProjectExample.docm");
    }
}
