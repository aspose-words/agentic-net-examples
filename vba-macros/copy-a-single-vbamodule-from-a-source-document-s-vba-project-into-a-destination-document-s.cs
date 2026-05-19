using System;
using Aspose.Words;
using Aspose.Words.Vba;

namespace VbaModuleCopyExample
{
    public class Program
    {
        public static void Main()
        {
            // Paths for the temporary source and destination documents.
            const string sourcePath = "Source.docm";
            const string destinationPath = "Destination.docm";

            // ---------- Create a source document with a VBA project and a single module ----------
            Document sourceDoc = new Document();

            // Create and assign a VBA project to the source document.
            VbaProject sourceProject = new VbaProject
            {
                Name = "SourceProject"
            };
            sourceDoc.VbaProject = sourceProject;

            // Create a VBA module, set its properties, and add it to the source project.
            VbaModule sourceModule = new VbaModule
            {
                Name = "MyModule",
                Type = VbaModuleType.ProceduralModule,
                SourceCode = "Sub HelloWorld()\n    MsgBox \"Hello from source\"\nEnd Sub"
            };
            sourceDoc.VbaProject.Modules.Add(sourceModule);

            // Save the source document in a macro‑enabled format.
            sourceDoc.Save(sourcePath, SaveFormat.Docm);

            // ---------- Create a destination document ----------
            Document destinationDoc = new Document();

            // Ensure the destination document has a VBA project.
            if (destinationDoc.VbaProject == null)
            {
                VbaProject destProject = new VbaProject
                {
                    Name = "DestinationProject"
                };
                destinationDoc.VbaProject = destProject;
            }

            // ---------- Copy the module from source to destination ----------
            // Retrieve the module by name from the source document.
            VbaModule moduleToCopy = sourceDoc.VbaProject.Modules["MyModule"];
            if (moduleToCopy != null)
            {
                // Clone the module to create an independent copy.
                VbaModule copiedModule = moduleToCopy.Clone();

                // Add the cloned module to the destination document's VBA project.
                destinationDoc.VbaProject.Modules.Add(copiedModule);
            }

            // Save the destination document, now containing the copied module.
            destinationDoc.Save(destinationPath, SaveFormat.Docm);
        }
    }
}
