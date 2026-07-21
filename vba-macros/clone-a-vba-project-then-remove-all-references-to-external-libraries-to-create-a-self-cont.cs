using System;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Create a blank document.
        Document sourceDoc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject vbaProject = new VbaProject();
        vbaProject.Name = "OriginalProject";
        sourceDoc.VbaProject = vbaProject;

        // Add a procedural module with sample VBA code.
        VbaModule module1 = new VbaModule();
        module1.Name = "Module1";
        module1.Type = VbaModuleType.ProceduralModule;
        module1.SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from VBA!""
End Sub";
        sourceDoc.VbaProject.Modules.Add(module1);

        // Add another module.
        VbaModule module2 = new VbaModule();
        module2.Name = "Module2";
        module2.Type = VbaModuleType.ProceduralModule;
        module2.SourceCode = @"
Function AddNumbers(a As Integer, b As Integer) As Integer
    AddNumbers = a + b
End Function";
        sourceDoc.VbaProject.Modules.Add(module2);

        // Save the source document (macro‑enabled).
        string sourcePath = "SourceMacro.docm";
        sourceDoc.Save(sourcePath);

        // Clone the VBA project from the source document.
        VbaProject clonedProject = sourceDoc.VbaProject.Clone();

        // Create a new document and assign the cloned VBA project.
        Document destDoc = new Document();
        destDoc.VbaProject = clonedProject;

        // Remove all references to external libraries from the cloned project.
        VbaReferenceCollection references = destDoc.VbaProject.References;
        for (int i = references.Count - 1; i >= 0; i--)
        {
            references.RemoveAt(i);
        }

        // Save the resulting document, which now contains a self‑contained macro set.
        string destPath = "ClonedNoReferences.docm";
        destDoc.Save(destPath);

        // Simple verification output.
        Console.WriteLine($"Source document saved to: {sourcePath}");
        Console.WriteLine($"Cloned document without references saved to: {destPath}");
        Console.WriteLine($"Cloned document has macros: {destDoc.HasMacros}");
        Console.WriteLine($"Number of references after removal: {destDoc.VbaProject.References.Count}");
    }
}
