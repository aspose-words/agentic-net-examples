using System;
using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Load the source document that contains the original VBA project.
        // Ensure the file exists; otherwise, the program will inform the user and exit gracefully.
        const string sourcePath = "Source.docm";
        if (!System.IO.File.Exists(sourcePath))
        {
            Console.WriteLine($"Source file \"{sourcePath}\" not found. The program will exit.");
            return;
        }

        Document sourceDoc = new Document(sourcePath);

        // Create a new (blank) destination document.
        Document destDoc = new Document();

        // Perform a deep clone of the VBA project from the source document, if it exists.
        VbaProject clonedProject = sourceDoc.VbaProject?.Clone();

        if (clonedProject != null)
        {
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
        }
        else
        {
            Console.WriteLine("The source document does not contain a VBA project. No VBA will be added to the destination document.");
        }

        // Save the resulting document; it now contains a self‑contained macro set (if any).
        const string destPath = "SelfContained.docm";
        destDoc.Save(destPath);
        Console.WriteLine($"Document saved to \"{destPath}\".");
    }
}
