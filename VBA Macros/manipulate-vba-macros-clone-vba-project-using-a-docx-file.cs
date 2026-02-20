using System;
using Aspose.Words;
using Aspose.Words.Vba;

class CloneVbaProjectExample
{
    static void Main()
    {
        // Load the source document that contains a VBA project.
        Document srcDoc = new Document("Input.docm");

        // Ensure the source document actually has macros.
        if (!srcDoc.HasMacros)
        {
            Console.WriteLine("Source document does not contain a VBA project.");
            return;
        }

        // Clone the VBA project from the source document.
        VbaProject clonedProject = srcDoc.VbaProject.Clone();

        // Create a new blank document.
        Document dstDoc = new Document();

        // Assign the cloned VBA project to the new document.
        dstDoc.VbaProject = clonedProject;

        // Save the new document with the cloned macros.
        dstDoc.Save("Cloned.docm");
    }
}
