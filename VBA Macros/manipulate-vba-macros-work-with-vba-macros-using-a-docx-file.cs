using System;
using Aspose.Words;
using Aspose.Words.Vba;

class VbaMacroExample
{
    static void Main()
    {
        // Load an existing DOCX file.
        Document doc = new Document("Input.docx");

        // Check whether the document already contains VBA macros.
        if (doc.HasMacros)
        {
            // Document already has a VBA project. Access it.
            VbaProject project = doc.VbaProject;

            // Output basic information about the existing VBA project.
            Console.WriteLine($"Existing VBA project name: {project.Name}");
            Console.WriteLine($"Modules count: {project.Modules.Count}");

            // Modify the source code of the first module (if any).
            if (project.Modules.Count > 0)
            {
                VbaModule firstModule = project.Modules[0];
                Console.WriteLine($"Modifying module: {firstModule.Name}");
                firstModule.SourceCode = @"Sub UpdatedMacro()
    MsgBox ""This macro has been updated by Aspose.Words.""
End Sub";
            }
        }
        else
        {
            // No VBA project present – create a new one.
            VbaProject newProject = new VbaProject
            {
                Name = "AsposeGeneratedProject",
                CodePage = 1252 // Western European (Windows) code page.
            };

            // Assign the new VBA project to the document.
            doc.VbaProject = newProject;

            // Create a new VBA module.
            VbaModule module = new VbaModule
            {
                Name = "AsposeModule",
                Type = VbaModuleType.ProceduralModule,
                SourceCode = @"Sub HelloWorld()
    MsgBox ""Hello from Aspose.Words VBA!""
End Sub"
            };

            // Add the module to the VBA project.
            doc.VbaProject.Modules.Add(module);
        }

        // Save the document as a macro‑enabled file (DOCM) to preserve the VBA project.
        doc.Save("Output.docm");
    }
}
