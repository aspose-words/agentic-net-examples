using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string sourcePath = Path.Combine(Directory.GetCurrentDirectory(), "Source.docm");
        string targetPath = Path.Combine(Directory.GetCurrentDirectory(), "Target.docm");

        // -------------------------------------------------
        // Step 1: Create a source macro‑enabled document.
        // -------------------------------------------------
        Document sourceDoc = new Document();

        // Create a new VBA project and give it a name.
        VbaProject sourceProject = new VbaProject();
        sourceProject.Name = "SourceProject";

        // Create a procedural module with simple VBA code.
        VbaModule module1 = new VbaModule();
        module1.Name = "Module1";
        module1.Type = VbaModuleType.ProceduralModule;
        module1.SourceCode = @"Sub HelloWorld()
    MsgBox ""Hello from source document!""
End Sub";

        // Add the module to the VBA project.
        sourceProject.Modules.Add(module1);

        // Assign the VBA project to the document.
        sourceDoc.VbaProject = sourceProject;

        // Save the source document as a macro‑enabled file.
        sourceDoc.Save(sourcePath);

        // -------------------------------------------------
        // Step 2: Load the source document and clone its VBA project.
        // -------------------------------------------------
        Document loadedSource = new Document(sourcePath);

        // Ensure the source document actually contains macros.
        if (!loadedSource.HasMacros)
        {
            throw new InvalidOperationException("Source document does not contain a VBA project.");
        }

        // Perform a deep clone of the VBA project.
        VbaProject clonedProject = loadedSource.VbaProject.Clone();

        // -------------------------------------------------
        // Step 3: Create a new document and attach the cloned VBA project.
        // -------------------------------------------------
        Document targetDoc = new Document();
        targetDoc.VbaProject = clonedProject;

        // Save the target document, which now contains the cloned macros.
        targetDoc.Save(targetPath);
    }
}
