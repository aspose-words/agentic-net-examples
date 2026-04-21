using System;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Create a blank document.
        Document originalDoc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject vbaProject = new VbaProject();
        vbaProject.Name = "OriginalProject";
        originalDoc.VbaProject = vbaProject;

        // Add a simple procedural module with some VBA code.
        VbaModule module = new VbaModule();
        module.Name = "SampleModule";
        module.Type = VbaModuleType.ProceduralModule;
        module.SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from VBA!""
End Sub
";
        originalDoc.VbaProject.Modules.Add(module);

        // Save the original macro‑enabled document.
        string originalPath = "Original.docm";
        originalDoc.Save(originalPath);

        // Clone the VBA project.
        VbaProject clonedProject = originalDoc.VbaProject.Clone();

        // Create a new document and assign the cloned project.
        Document clonedDoc = new Document();
        clonedDoc.VbaProject = clonedProject;

        // Remove all external references from the cloned VBA project.
        VbaReferenceCollection references = clonedDoc.VbaProject.References;
        for (int i = references.Count - 1; i >= 0; i--)
        {
            // Remove each reference by index.
            references.RemoveAt(i);
        }

        // Save the cloned document which now has no external references.
        string clonedPath = "Cloned_NoReferences.docm";
        clonedDoc.Save(clonedPath);

        // Simple verification output.
        Console.WriteLine($"Original document saved to: {originalPath}");
        Console.WriteLine($"Cloned document without references saved to: {clonedPath}");
        Console.WriteLine($"Original has references: {originalDoc.VbaProject.References.Count}");
        Console.WriteLine($"Cloned has references: {clonedDoc.VbaProject.References.Count}");
    }
}
