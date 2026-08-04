using System;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Paths for the original and updated macro-enabled documents.
        const string originalPath = "MacroDocument.docm";
        const string updatedPath = "MacroDocumentUpdated.docm";

        // 1. Create a blank document.
        Document doc = new Document();

        // 2. Create a new VBA project and assign it to the document.
        VbaProject vbaProject = new VbaProject();
        vbaProject.Name = "SampleProject";
        doc.VbaProject = vbaProject;

        // 3. Create a VBA module with sample macro code.
        VbaModule module = new VbaModule();
        module.Name = "SampleModule";
        module.Type = VbaModuleType.ProceduralModule;
        module.SourceCode = @"
Sub SampleMacro()
    Dim x As Integer
    x = 1 / 0   ' This will cause a division by zero error
    MsgBox ""Result: "" & x
End Sub";

        // 4. Add the module to the VBA project.
        doc.VbaProject.Modules.Add(module);

        // 5. Save the document as a macro‑enabled file.
        doc.Save(originalPath);

        // 6. Load the saved document (could also continue with the same instance).
        Document loadedDoc = new Document(originalPath);

        // 7. Verify the document contains macros before attempting an update.
        if (loadedDoc.HasMacros && loadedDoc.VbaProject != null)
        {
            // Retrieve the module by name.
            VbaModule targetModule = loadedDoc.VbaProject.Modules["SampleModule"];
            if (targetModule != null)
            {
                // Guard against null source code.
                string source = targetModule.SourceCode ?? string.Empty;

                // Define the error‑handling statement.
                const string errorHandler = "On Error Resume Next";

                // Prepend the error handler if it is not already present.
                if (!source.Contains(errorHandler))
                {
                    source = errorHandler + Environment.NewLine + source;
                    targetModule.SourceCode = source;
                }
            }
        }

        // 8. Save the updated document.
        loadedDoc.Save(updatedPath);
    }
}
