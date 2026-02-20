using System;
using Aspose.Words;
using Aspose.Words.Vba;

namespace VbaModuleCloneExample
{
    class Program
    {
        static void Main()
        {
            // Load the source document that already contains a VBA project.
            Document doc = new Document("InputDocument.docm");

            // Access the VBA project of the document.
            VbaProject vbaProject = doc.VbaProject;

            // Ensure there is at least one module to clone.
            if (vbaProject.Modules.Count == 0)
            {
                Console.WriteLine("The document does not contain any VBA modules.");
                return;
            }

            // Get the first module (or any specific module by index or name).
            VbaModule originalModule = vbaProject.Modules[0];

            // Clone the module. The Clone method creates a deep copy of the module.
            VbaModule clonedModule = originalModule.Clone();

            // Optionally give the cloned module a new name to avoid name conflicts.
            clonedModule.Name = originalModule.Name + "_Copy";

            // Add the cloned module back into the VBA project.
            vbaProject.Modules.Add(clonedModule);

            // Save the document with the new cloned module.
            doc.Save("OutputDocument.docm");
        }
    }
}
