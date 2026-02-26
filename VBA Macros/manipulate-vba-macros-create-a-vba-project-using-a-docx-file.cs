using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Load an existing DOCX file.
        Document doc = new Document("Input.docx");

        // Create a new VBA project and assign a name.
        VbaProject project = new VbaProject();
        project.Name = "MyProject";

        // Attach the VBA project to the document.
        doc.VbaProject = project;

        // Create a new procedural module that will contain the macro code.
        VbaModule module = new VbaModule();
        module.Name = "MyModule";
        module.Type = VbaModuleType.ProceduralModule;
        module.SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from VBA!""
End Sub";

        // Add the module to the VBA project's module collection.
        doc.VbaProject.Modules.Add(module);

        // Save the document as a macro‑enabled file (.docm).
        doc.Save("Output.docm");
    }
}
