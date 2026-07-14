using System;
using Aspose.Words;
using Aspose.Words.Vba;

namespace VbaCloneAndStripReferences
{
    public class Program
    {
        public static void Main()
        {
            // Create a blank document.
            Document originalDoc = new Document();

            // Create a new VBA project and assign a name.
            VbaProject vbaProject = new VbaProject { Name = "OriginalProject" };
            originalDoc.VbaProject = vbaProject;

            // Add a simple procedural module with sample macro code.
            VbaModule module = new VbaModule
            {
                Name = "SampleModule",
                Type = VbaModuleType.ProceduralModule,
                SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from VBA!""
End Sub"
            };
            originalDoc.VbaProject.Modules.Add(module);

            // Save the original macro‑enabled document.
            const string originalPath = "Original.docm";
            originalDoc.Save(originalPath);

            // Clone the VBA project.
            VbaProject clonedProject = originalDoc.VbaProject.Clone();

            // Create a new document and assign the cloned project.
            Document clonedDoc = new Document();
            clonedDoc.VbaProject = clonedProject;

            // Remove all references from the cloned project to make it self‑contained.
            // Iterate backwards because we are modifying the collection while iterating.
            for (int i = clonedDoc.VbaProject.References.Count - 1; i >= 0; i--)
            {
                clonedDoc.VbaProject.References.RemoveAt(i);
            }

            // Save the cloned, self‑contained macro‑enabled document.
            const string clonedPath = "ClonedSelfContained.docm";
            clonedDoc.Save(clonedPath);

            // Simple verification output.
            Console.WriteLine($"Original document saved to: {originalPath}");
            Console.WriteLine($"Cloned document saved to: {clonedPath}");
            Console.WriteLine($"Cloned document has {clonedDoc.VbaProject.References.Count} VBA references (expected 0).");
        }
    }
}
