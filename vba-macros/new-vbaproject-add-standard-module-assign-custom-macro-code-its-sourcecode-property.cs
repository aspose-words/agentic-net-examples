using System;
using Aspose.Words;
using Aspose.Words.Vba;

class VbaProjectExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialise a new VBA project and assign a name.
        VbaProject vbaProject = new VbaProject
        {
            Name = "MyVbaProject"
        };
        doc.VbaProject = vbaProject;

        // Create a standard procedural module.
        VbaModule vbaModule = new VbaModule
        {
            Name = "MyStandardModule",
            Type = VbaModuleType.ProceduralModule,
            // Assign custom macro source code to the module.
            SourceCode = @"Sub MyMacro()
    MsgBox ""Hello from VBA macro!""
End Sub"
        };

        // Add the module to the VBA project's module collection.
        vbaProject.Modules.Add(vbaModule);

        // Save the document as a macro‑enabled file.
        doc.Save("VbaProjectWithMacro.docm");
    }
}
