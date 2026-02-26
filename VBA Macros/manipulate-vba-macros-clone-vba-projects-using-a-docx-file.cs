using System;
using Aspose.Words;
using Aspose.Words.Vba;

class VbaProjectCloner
{
    static void Main()
    {
        // Path to the folder that contains the source document.
        string dataDir = @"C:\Data\";

        // Load the source document that contains VBA macros.
        // The document must be a macro-enabled format (e.g., .docm) for the VbaProject to be present.
        Document sourceDoc = new Document(dataDir + "Source.docm");

        // Clone the VBA project from the source document.
        VbaProject clonedProject = sourceDoc.VbaProject.Clone();

        // Create a new empty document that will receive the cloned VBA project.
        Document destinationDoc = new Document();

        // Assign the cloned VBA project to the new document.
        destinationDoc.VbaProject = clonedProject;

        // Save the destination document as a macro‑enabled file.
        destinationDoc.Save(dataDir + "Cloned.docm");
    }
}
