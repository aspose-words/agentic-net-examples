using System;
using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Load an existing DOCX file.
        Document doc = new Document("Input.docx");

        // If the document does not already contain a VBA project, create one.
        if (!doc.HasMacros)
        {
            VbaProject project = new VbaProject();
            project.Name = "MyProject";
            doc.VbaProject = project;
        }

        // Create a new procedural VBA module.
        VbaModule module = new VbaModule
        {
            Name = "MyMacroModule",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from Aspose.Words!""
End Sub
"
        };

        // Add the module to the document's VBA project.
        doc.VbaProject.Modules.Add(module);

        // Save the document as a macro‑enabled file.
        doc.Save("Output.docm");
    }
}
