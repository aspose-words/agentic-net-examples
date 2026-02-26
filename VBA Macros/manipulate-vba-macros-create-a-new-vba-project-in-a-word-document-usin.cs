using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Load an existing DOCX file.
        Document doc = new Document("Input.docx");

        // Create a new VBA project and assign it to the document.
        VbaProject project = new VbaProject();
        project.Name = "MyProject";
        doc.VbaProject = project;

        // Create a new procedural module with VBA source code.
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
        doc.Save("Output.docm");
    }
}
