using System;
using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Load the source document (DOCX). It may not contain macros, so we create a VBA project if needed.
        Document doc = new Document("Input.docx");

        // Ensure the document has a VBA project to hold modules.
        if (doc.VbaProject == null)
            doc.VbaProject = new VbaProject();

        // Create an original VBA module.
        VbaModule original = new VbaModule
        {
            Name = "OriginalModule",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = "Sub Hello()\n    MsgBox \"Hello\"\nEnd Sub"
        };

        // Add the original module to the project's module collection.
        doc.VbaProject.Modules.Add(original);

        // Clone the original module.
        VbaModule cloned = original.Clone();

        // Give the cloned module a distinct name.
        cloned.Name = "ClonedModule";

        // Add the cloned module to the same VBA project.
        doc.VbaProject.Modules.Add(cloned);

        // Save the document as a macro‑enabled file so the VBA project is preserved.
        doc.Save("Output.docm");
    }
}
