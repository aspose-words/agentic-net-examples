using System;
using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Load the source document that contains the original VBA project.
        Document sourceDoc = new Document("Source.docm");

        // Create a new (blank) destination document.
        Document destDoc = new Document();

        // Perform a deep clone of the VBA project from the source document.
        VbaProject clonedProject = sourceDoc.VbaProject.Clone();

        // Remove all references from the cloned project to make it self‑contained.
        VbaReferenceCollection references = clonedProject.References;
        for (int i = references.Count - 1; i >= 0; i--)
        {
            // If you only want to keep project references, uncomment the condition below.
            // if (references[i].Type != VbaReferenceType.Project)
            // {
                references.RemoveAt(i);
            // }
        }

        // Assign the cleaned, cloned VBA project to the destination document.
        destDoc.VbaProject = clonedProject;

        // Save the resulting document; it now contains a self‑contained macro set.
        destDoc.Save("SelfContained.docm");
    }
}
