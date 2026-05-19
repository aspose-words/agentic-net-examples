using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Folder for output files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a macro‑enabled document with a VBA project and a module.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject sourceProject = new VbaProject();
        sourceProject.Name = "SourceProject";
        sourceDoc.VbaProject = sourceProject;

        // Add a simple procedural module with some VBA code.
        VbaModule module = new VbaModule();
        module.Name = "SampleModule";
        module.Type = VbaModuleType.ProceduralModule;
        module.SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from Aspose.Words!""
End Sub
";
        sourceDoc.VbaProject.Modules.Add(module);

        // Save the source document (macro‑enabled format).
        string sourcePath = Path.Combine(outputDir, "Source.docm");
        sourceDoc.Save(sourcePath);

        // ---------------------------------------------------------------
        // 2. Clone the VBA project from the source document.
        // ---------------------------------------------------------------
        VbaProject clonedProject = sourceDoc.VbaProject.Clone();

        // Create a new document and assign the cloned VBA project.
        Document destDoc = new Document();
        destDoc.VbaProject = clonedProject;

        // ---------------------------------------------------------------
        // 3. Remove all external references from the cloned VBA project.
        // ---------------------------------------------------------------
        // The References collection may be empty if no references were added,
        // but the loop safely removes any that exist.
        VbaReferenceCollection references = destDoc.VbaProject.References;
        for (int i = references.Count - 1; i >= 0; i--)
        {
            references.RemoveAt(i);
        }

        // Save the resulting document, which now contains a self‑contained macro set.
        string destPath = Path.Combine(outputDir, "ClonedNoRefs.docm");
        destDoc.Save(destPath);
    }
}
