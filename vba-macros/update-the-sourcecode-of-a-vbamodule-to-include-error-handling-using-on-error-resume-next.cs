using System;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Paths for the original and updated macro-enabled documents.
        const string originalPath = "MacroOriginal.docm";
        const string updatedPath = "MacroUpdated.docm";

        // 1. Create a blank document.
        Document doc = new Document();

        // 2. Create a VBA project and assign it to the document.
        VbaProject project = new VbaProject
        {
            Name = "SampleProject"
        };
        doc.VbaProject = project;

        // 3. Create a VBA module with a simple macro.
        VbaModule module = new VbaModule
        {
            Name = "SampleModule",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"Sub HelloWorld()
    MsgBox ""Hello, World!""
End Sub"
        };

        // 4. Add the module to the VBA project.
        doc.VbaProject.Modules.Add(module);

        // 5. Save the document that now contains the macro.
        doc.Save(originalPath);

        // 6. Load the saved document to demonstrate updating the module source.
        Document loadedDoc = new Document(originalPath);

        // 7. Verify that the document has macros and at least one module.
        if (loadedDoc.HasMacros && loadedDoc.VbaProject != null && loadedDoc.VbaProject.Modules.Count > 0)
        {
            // Retrieve the first module (or use a name lookup if preferred).
            VbaModule targetModule = loadedDoc.VbaProject.Modules[0];

            // Guard against null source code.
            string currentCode = targetModule.SourceCode ?? string.Empty;

            // Define the error‑handling statement.
            const string errorHandling = "On Error Resume Next";

            // Prepend the error‑handling line if it is not already present.
            if (!currentCode.Contains(errorHandling))
            {
                targetModule.SourceCode = errorHandling + "\r\n" + currentCode;
            }
        }

        // 8. Save the document with the updated VBA source.
        loadedDoc.Save(updatedPath);
    }
}
