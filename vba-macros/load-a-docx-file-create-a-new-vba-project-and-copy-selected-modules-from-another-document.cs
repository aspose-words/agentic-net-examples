using System;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        const string inputDocxPath = "Input.docx";
        const string sourceMacroDocPath = "Source.docm";
        const string outputDocmPath = "Result.docm";

        // -----------------------------------------------------------------
        // Step 1: Create a plain DOCX file (no macros) and save it.
        // -----------------------------------------------------------------
        Document plainDoc = new Document();
        plainDoc.Save(inputDocxPath); // Saved as DOCX by default.

        // -----------------------------------------------------------------
        // Step 2: Create a macro-enabled document that will serve as the
        // source of VBA modules. Add a VBA project with two sample modules.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();

        // Create a new VBA project.
        VbaProject sourceProject = new VbaProject
        {
            Name = "SourceProject"
        };
        sourceDoc.VbaProject = sourceProject;

        // First module.
        VbaModule module1 = new VbaModule
        {
            Name = "Module1",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from Module1!""
End Sub"
        };
        sourceProject.Modules.Add(module1);

        // Second module.
        VbaModule module2 = new VbaModule
        {
            Name = "Module2",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"
Sub GoodbyeWorld()
    MsgBox ""Goodbye from Module2!""
End Sub"
        };
        sourceProject.Modules.Add(module2);

        // Save the source document as a macro-enabled file.
        sourceDoc.Save(sourceMacroDocPath); // .docm inferred from extension.

        // -----------------------------------------------------------------
        // Step 3: Load the plain DOCX file.
        // -----------------------------------------------------------------
        Document targetDoc = new Document(inputDocxPath);

        // Ensure the document has a VBA project; create one if missing.
        if (targetDoc.VbaProject == null)
        {
            VbaProject newProject = new VbaProject
            {
                Name = "TargetProject"
            };
            targetDoc.VbaProject = newProject;
        }

        // -----------------------------------------------------------------
        // Step 4: Load the source macro document and copy selected modules.
        // -----------------------------------------------------------------
        Document sourceMacroDoc = new Document(sourceMacroDocPath);
        VbaProject sourceVbaProject = sourceMacroDoc.VbaProject;

        // Example: copy modules whose names start with "Module".
        foreach (VbaModule srcModule in sourceVbaProject.Modules)
        {
            if (srcModule.Name != null && srcModule.Name.StartsWith("Module"))
            {
                // Clone the module to avoid reference issues.
                VbaModule clonedModule = srcModule.Clone();

                // Ensure source code is not null.
                if (clonedModule.SourceCode == null)
                    clonedModule.SourceCode = string.Empty;

                // Add the cloned module to the target document's VBA project.
                targetDoc.VbaProject.Modules.Add(clonedModule);
            }
        }

        // -----------------------------------------------------------------
        // Step 5: Save the modified document as a macro-enabled file.
        // -----------------------------------------------------------------
        targetDoc.Save(outputDocmPath); // Saved as .docm because of extension.

        // Optional: indicate completion.
        Console.WriteLine("Macro modules copied and document saved as " + outputDocmPath);
    }
}
