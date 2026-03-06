using System;
using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Path to the existing DOCX file.
        string inputPath = @"C:\Docs\input.docx";

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // Create a new VBA project.
        VbaProject project = new VbaProject();
        project.Name = "MyVbaProject";

        // Assign the VBA project to the document.
        doc.VbaProject = project;

        // Create a new VBA module.
        VbaModule module = new VbaModule();
        module.Name = "MyModule";
        module.Type = VbaModuleType.ProceduralModule;
        module.SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from VBA!""
End Sub";

        // Add the module to the VBA project.
        doc.VbaProject.Modules.Add(module);

        // Save the document as a macro-enabled file.
        string outputPath = @"C:\Docs\output.docm";
        doc.Save(outputPath);
    }
}
