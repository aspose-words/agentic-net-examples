using System;
using Aspose.Words;
using Aspose.Words.Vba;

class ModifyVbaModule
{
    static void Main()
    {
        // Path to the source DOCX file (must exist)
        string inputPath = @"C:\Docs\input.docx";

        // Path where the modified document will be saved (DOCM to keep macros)
        string outputPath = @"C:\Docs\output.docm";

        // Load the document
        Document doc = new Document(inputPath);

        // Ensure the document has a VBA project; create one if it doesn't
        if (doc.VbaProject == null)
        {
            doc.VbaProject = new VbaProject
            {
                Name = "MyVbaProject"
            };
        }

        VbaProject vbaProject = doc.VbaProject;

        // Retrieve an existing module or create a new one
        VbaModule vbaModule;
        if (vbaProject.Modules.Count > 0)
        {
            // Use the first module in the collection
            vbaModule = vbaProject.Modules[0];
        }
        else
        {
            // Create a new procedural module
            vbaModule = new VbaModule
            {
                Name = "Module1",
                Type = VbaModuleType.ProceduralModule
            };
            vbaProject.Modules.Add(vbaModule);
        }

        // Replace the source code of the selected module
        vbaModule.SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from Aspose.Words!""
End Sub";

        // Save the document as a macro‑enabled file (DOCM)
        doc.Save(outputPath, SaveFormat.Docm);
    }
}
