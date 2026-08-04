using System;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Path for the temporary files.
        string sourcePath = "source.docm";
        string resultPath = "cloned_no_refs.docm";

        // -------------------------------------------------
        // 1. Create a macro‑enabled document with a VBA project.
        // -------------------------------------------------
        Document sourceDoc = new Document();

        // Create a new VBA project and assign a name.
        VbaProject vbaProject = new VbaProject { Name = "OriginalProject" };

        // Create a simple procedural module with some VBA code.
        VbaModule module = new VbaModule
        {
            Name = "SampleModule",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello, World!""
End Sub"
        };

        // Add the module to the project.
        vbaProject.Modules.Add(module);

        // Attach the VBA project to the document.
        sourceDoc.VbaProject = vbaProject;

        // Save the document in a macro‑enabled format.
        sourceDoc.Save(sourcePath);

        // -------------------------------------------------
        // 2. Load the document and clone its VBA project.
        // -------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        // Ensure the document actually contains macros.
        if (!loadedDoc.HasMacros || loadedDoc.VbaProject == null)
        {
            Console.WriteLine("The source document does not contain a VBA project.");
            return;
        }

        // Clone the VBA project.
        VbaProject clonedProject = loadedDoc.VbaProject.Clone();

        // -------------------------------------------------
        // 3. Remove all external references from the cloned project.
        // -------------------------------------------------
        // The References collection may be empty; iterate backwards to safely remove items.
        for (int i = clonedProject.References.Count - 1; i >= 0; i--)
        {
            // Remove the reference at the current index.
            clonedProject.References.RemoveAt(i);
        }

        // -------------------------------------------------
        // 4. Create a new document and assign the cleaned VBA project.
        // -------------------------------------------------
        Document resultDoc = new Document();
        resultDoc.VbaProject = clonedProject;

        // Save the resulting document; it now contains the same macros but no external references.
        resultDoc.Save(resultPath);
    }
}
