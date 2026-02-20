using System;
using Aspose.Words;
using Aspose.Words.Vba;

class CreateVbaProject
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a new VBA project and assign a name.
        VbaProject project = new VbaProject
        {
            Name = "MyVbaProject"
        };

        // Attach the VBA project to the document.
        doc.VbaProject = project;

        // Create a new procedural module that will hold the macro code.
        VbaModule module = new VbaModule
        {
            Name = "MyMacroModule",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from VBA!""
End Sub"
        };

        // Add the module to the VBA project's module collection.
        doc.VbaProject.Modules.Add(module);

        // Save the document in a macro-enabled format (DOCM).
        doc.Save("VbaProjectCreated.docm", SaveFormat.Docm);
    }
}
