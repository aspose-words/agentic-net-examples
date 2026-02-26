using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Load the source document that already contains a VBA project.
        Document srcDoc = new Document("SourceWithMacros.docm");

        // Create a new empty document that will receive the cloned VBA project.
        Document destDoc = new Document();

        // Clone the VBA project from the source document.
        VbaProject clonedProject = srcDoc.VbaProject.Clone();

        // Assign the cloned project to the destination document.
        destDoc.VbaProject = clonedProject;

        // Save the destination document as a macro‑enabled file.
        destDoc.Save("ClonedVbaProject.docm");
    }
}
