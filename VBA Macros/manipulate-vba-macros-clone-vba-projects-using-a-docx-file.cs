using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Load the source document that already contains VBA macros.
        Document sourceDoc = new Document("SourceWithMacros.docm");

        // Clone the entire VBA project (all modules, references, etc.).
        VbaProject clonedProject = sourceDoc.VbaProject.Clone();

        // Create a new blank document.
        Document targetDoc = new Document();

        // Attach the cloned VBA project to the new document.
        targetDoc.VbaProject = clonedProject;

        // Save the new document; it will now contain the cloned macros.
        targetDoc.Save("ClonedMacros.docm");
    }
}
