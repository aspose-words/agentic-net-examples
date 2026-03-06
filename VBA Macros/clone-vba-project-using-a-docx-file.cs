using System;
using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Load the source document that contains a VBA project.
        Document srcDoc = new Document("Source.docm");

        // Create a new blank destination document.
        Document destDoc = new Document();

        // Clone the entire VBA project from the source document.
        VbaProject clonedProject = srcDoc.VbaProject.Clone();
        destDoc.VbaProject = clonedProject;

        // The clone also copies all modules, so a module with the same name may already exist.
        // Remove the duplicated module and replace it with a fresh clone from the source.
        VbaModule oldModule = destDoc.VbaProject.Modules["Module1"];
        if (oldModule != null)
        {
            destDoc.VbaProject.Modules.Remove(oldModule);
        }

        VbaModule newModule = srcDoc.VbaProject.Modules["Module1"].Clone();
        destDoc.VbaProject.Modules.Add(newModule);

        // Save the destination document with the cloned VBA project.
        destDoc.Save("ClonedVbaProject.docm");
    }
}
