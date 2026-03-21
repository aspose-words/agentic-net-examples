using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

class CloneVbaProjectExample
{
    static void Main()
    {
        const string sourcePath = "VBA project.docm";
        const string destPath = "VbaProject.CloneVbaProject.docm";

        // If the source document does not exist, create a minimal macro‑enabled document.
        Document srcDoc;
        if (File.Exists(sourcePath))
        {
            srcDoc = new Document(sourcePath);
        }
        else
        {
            Console.WriteLine($"Source file \"{sourcePath}\" not found. Creating a temporary macro‑enabled document.");

            // Create a new blank document and enable the VBA project container.
            srcDoc = new Document();
            srcDoc.Save(sourcePath, SaveFormat.Docm);
            srcDoc = new Document(sourcePath);
        }

        // Create a new empty document that will receive the cloned VBA project.
        Document destDoc = new Document();

        // Ensure the source document actually contains a VBA project before cloning.
        if (srcDoc.VbaProject != null && srcDoc.VbaProject.Modules.Count > 0)
        {
            // Clone the entire VBA project (includes references).
            VbaProject clonedProject = srcDoc.VbaProject.Clone();

            // Assign the cloned project to the destination document.
            destDoc.VbaProject = clonedProject;

            // Remove any default modules that may have been added automatically.
            while (destDoc.VbaProject.Modules.Count > 0)
            {
                destDoc.VbaProject.Modules.Remove(destDoc.VbaProject.Modules[0]);
            }

            // Add cloned modules from the source document preserving their original order.
            foreach (VbaModule srcModule in srcDoc.VbaProject.Modules)
            {
                VbaModule clonedModule = srcModule.Clone();
                destDoc.VbaProject.Modules.Add(clonedModule);
            }
        }
        else
        {
            Console.WriteLine("Source document does not contain any VBA modules. The destination document will be saved without VBA.");
        }

        // Save the document with the duplicated VBA project, preserving module order and references.
        destDoc.Save(destPath);
        Console.WriteLine($"Document saved to \"{destPath}\".");
    }
}
