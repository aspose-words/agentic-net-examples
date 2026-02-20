using System;
using Aspose.Words;
using Aspose.Words.Vba;

namespace VbaMacroManipulation
{
    class Program
    {
        static void Main()
        {
            // Load an existing macro-enabled document (DOCM).
            Document doc = new Document("InputDocument.docm");

            // Ensure the document actually contains a VBA project.
            if (!doc.HasMacros)
            {
                Console.WriteLine("The document does not contain any VBA macros.");
                return;
            }

            // Access the VBA project.
            VbaProject vbaProject = doc.VbaProject;

            // Access a specific module by name (e.g., "Module1").
            // If the module does not exist, you could add a new one instead.
            VbaModule targetModule = vbaProject.Modules["Module1"];

            // Replace the source code of the module with new VBA code.
            targetModule.SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from Aspose.Words!""
End Sub
";

            // Optionally, you can change the module name or type.
            // targetModule.Name = "NewModuleName";
            // targetModule.Type = VbaModuleType.ProceduralModule;

            // Save the document back as a macro-enabled file.
            doc.Save("OutputDocument.docm");
        }
    }
}
