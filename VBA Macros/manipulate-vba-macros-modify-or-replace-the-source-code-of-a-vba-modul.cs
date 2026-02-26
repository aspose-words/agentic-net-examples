using System;
using Aspose.Words;
using Aspose.Words.Vba;

class VbaMacroModifier
{
    static void Main()
    {
        // Path to the input DOCX file (may not contain macros).
        string inputPath = @"C:\Docs\Input.docx";

        // Load the document.
        Document doc = new Document(inputPath);

        // Ensure the document has a VBA project; create one if it does not exist.
        if (doc.VbaProject == null)
        {
            doc.VbaProject = new VbaProject
            {
                Name = "MyVbaProject"
            };
        }

        // Access the VBA project.
        VbaProject vbaProject = doc.VbaProject;

        // Retrieve the collection of VBA modules.
        VbaModuleCollection modules = vbaProject.Modules;

        // If there is no module named "Module1", create a new procedural module.
        VbaModule targetModule = modules["Module1"];
        if (targetModule == null)
        {
            targetModule = new VbaModule
            {
                Name = "Module1",
                Type = VbaModuleType.ProceduralModule,
                SourceCode = string.Empty
            };
            modules.Add(targetModule);
        }

        // Replace the source code of the target module.
        targetModule.SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello, World!""
End Sub
";

        // Save the document as a macro-enabled file (DOCM) to preserve the VBA project.
        string outputPath = @"C:\Docs\Output.docm";
        doc.Save(outputPath);
    }
}
