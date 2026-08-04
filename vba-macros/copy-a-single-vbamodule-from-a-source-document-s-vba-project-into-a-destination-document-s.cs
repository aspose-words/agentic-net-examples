using System;
using Aspose.Words;
using Aspose.Words.Vba;

namespace VbaModuleCopyExample
{
    public class Program
    {
        public static void Main()
        {
            // Paths for the sample files.
            string sourcePath = "Source.docm";
            string destinationPath = "Destination.docm";

            // ---------- Create a source document with a VBA project and a single module ----------
            Document sourceDoc = new Document();

            // Create and assign a VBA project.
            VbaProject sourceProject = new VbaProject
            {
                Name = "SourceProject"
            };
            sourceDoc.VbaProject = sourceProject;

            // Create a VBA module with some simple macro code.
            VbaModule sourceModule = new VbaModule
            {
                Name = "SampleModule",
                Type = VbaModuleType.ProceduralModule,
                SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from the source module!""
End Sub"
            };

            // Add the module to the source project's collection.
            sourceProject.Modules.Add(sourceModule);

            // Save the source document as a macro‑enabled file.
            sourceDoc.Save(sourcePath);

            // ---------- Create a destination document ----------
            Document destDoc = new Document();

            // Ensure the destination document has a VBA project.
            if (destDoc.VbaProject == null)
            {
                VbaProject destProject = new VbaProject
                {
                    Name = "DestinationProject"
                };
                destDoc.VbaProject = destProject;
            }

            // ---------- Copy the module from source to destination ----------
            // Retrieve the module to copy (by name).
            VbaModule moduleToCopy = sourceDoc.VbaProject.Modules["SampleModule"];
            if (moduleToCopy != null)
            {
                // Clone the module to create an independent copy.
                VbaModule copiedModule = moduleToCopy.Clone();

                // Add the cloned module to the destination project's collection.
                destDoc.VbaProject.Modules.Add(copiedModule);
            }

            // Save the destination document as a macro‑enabled file.
            destDoc.Save(destinationPath);

            // Simple verification output.
            Console.WriteLine($"Module '{sourceModule.Name}' copied from '{sourcePath}' to '{destinationPath}'.");
        }
    }
}
