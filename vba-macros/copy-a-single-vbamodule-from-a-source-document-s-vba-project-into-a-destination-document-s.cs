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

            // -------------------------------------------------
            // Create a source document with a VBA project and a single module.
            // -------------------------------------------------
            Document sourceDoc = new Document();

            // Create and assign a new VBA project.
            VbaProject sourceProject = new VbaProject
            {
                Name = "SourceProject"
            };
            sourceDoc.VbaProject = sourceProject;

            // Create a VBA module, set its properties, and add it to the project.
            VbaModule sourceModule = new VbaModule
            {
                Name = "MyModule",
                Type = VbaModuleType.ProceduralModule,
                SourceCode = "Sub HelloWorld()\n    MsgBox \"Hello from source\"\nEnd Sub"
            };
            sourceProject.Modules.Add(sourceModule);

            // Save the source document in a macro‑enabled format.
            sourceDoc.Save(sourcePath);

            // -------------------------------------------------
            // Load the source document (simulating an existing file).
            // -------------------------------------------------
            Document src = new Document(sourcePath);

            // -------------------------------------------------
            // Create a destination document and its VBA project.
            // -------------------------------------------------
            Document dest = new Document();

            VbaProject destProject = new VbaProject
            {
                Name = "DestinationProject"
            };
            dest.VbaProject = destProject;

            // -------------------------------------------------
            // Copy the specified module from the source to the destination.
            // -------------------------------------------------
            VbaModule moduleToCopy = src.VbaProject?.Modules["MyModule"];
            if (moduleToCopy != null)
            {
                // Clone creates a deep copy of the module.
                VbaModule copiedModule = moduleToCopy.Clone();

                // Add the cloned module to the destination project's collection.
                destProject.Modules.Add(copiedModule);
            }

            // -------------------------------------------------
            // Save the destination document with the copied module.
            // -------------------------------------------------
            dest.Save(destinationPath);
        }
    }
}
