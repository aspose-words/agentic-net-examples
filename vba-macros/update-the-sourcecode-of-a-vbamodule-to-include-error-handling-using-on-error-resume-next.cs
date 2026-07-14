using System;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject vbaProject = new VbaProject();
        vbaProject.Name = "SampleProject";
        doc.VbaProject = vbaProject;

        // Create a new VBA module with simple macro code (no error handling).
        VbaModule vbaModule = new VbaModule();
        vbaModule.Name = "TestModule";
        vbaModule.Type = VbaModuleType.ProceduralModule;
        vbaModule.SourceCode = "Sub TestMacro()\n    MsgBox \"Hello\"\nEnd Sub";

        // Add the module to the VBA project.
        doc.VbaProject.Modules.Add(vbaModule);

        // Save the document as a macro-enabled file.
        string initialPath = "MacroWithoutErrorHandling.docm";
        doc.Save(initialPath);

        // Load the saved document (optional, can continue using the same instance).
        Document loadedDoc = new Document(initialPath);

        // Ensure the document has a VBA project and the target module exists.
        if (loadedDoc.HasMacros && loadedDoc.VbaProject != null)
        {
            VbaModule targetModule = loadedDoc.VbaProject.Modules["TestModule"];
            if (targetModule != null)
            {
                // Guard against null source code.
                string source = targetModule.SourceCode ?? string.Empty;

                // Add "On Error Resume Next" at the beginning if it's not already present.
                if (!source.StartsWith("On Error Resume Next", StringComparison.OrdinalIgnoreCase))
                {
                    targetModule.SourceCode = "On Error Resume Next\n" + source;
                }
            }
        }

        // Save the updated document.
        string updatedPath = "MacroWithErrorHandling.docm";
        loadedDoc.Save(updatedPath);

        // Indicate completion.
        Console.WriteLine("VBA module source code updated with error handling.");
    }
}
