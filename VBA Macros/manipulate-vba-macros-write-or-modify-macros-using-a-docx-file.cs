using System;
using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Load an existing document (DOCX or DOCM).
        string inputPath = "Input.docx";
        Document doc = new Document(inputPath);

        // Determine whether the document already contains VBA macros.
        if (doc.HasMacros)
        {
            // Document has macros – modify the source code of the first module.
            VbaModule firstModule = doc.VbaProject.Modules[0];
            firstModule.SourceCode = @"Sub HelloWorld()
    MsgBox ""Hello from modified macro!""
End Sub";
        }
        else
        {
            // Document has no macros – create a new VBA project and add a module.
            VbaProject project = new VbaProject();
            project.Name = "AsposeProject";
            doc.VbaProject = project;

            VbaModule module = new VbaModule();
            module.Name = "HelloModule";
            module.Type = VbaModuleType.ProceduralModule;
            module.SourceCode = @"Sub HelloWorld()
    MsgBox ""Hello from new macro!""
End Sub";

            doc.VbaProject.Modules.Add(module);
        }

        // Save the document as a macro‑enabled file.
        string outputPath = "Output.docm";
        doc.Save(outputPath);
    }
}
