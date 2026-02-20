using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Load the source document that contains a VBA project.
        Document sourceDoc = new Document("Source.docm");

        // Clone the VBA project from the source document.
        VbaProject clonedProject = sourceDoc.VbaProject.Clone();

        // Create a new blank document.
        Document targetDoc = new Document();

        // Assign the cloned VBA project to the new document.
        targetDoc.VbaProject = clonedProject;

        // Save the new document with the cloned macros.
        targetDoc.Save("ClonedProject.docm");
    }
}
