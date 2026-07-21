using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

namespace VbaModuleCopyExample
{
    public class Program
    {
        public static void Main()
        {
            // Define file names for the source and destination documents.
            string sourcePath = Path.Combine(Directory.GetCurrentDirectory(), "Source.docm");
            string destPath = Path.Combine(Directory.GetCurrentDirectory(), "Destination.docm");

            // ---------- Create the source document with a VBA project and a single module ----------
            Document sourceDoc = new Document();

            // Create a new VBA project and assign it to the source document.
            VbaProject sourceProject = new VbaProject
            {
                Name = "SourceProject"
            };
            sourceDoc.VbaProject = sourceProject;

            // Create a VBA module, set its properties, and add it to the source project.
            VbaModule sourceModule = new VbaModule
            {
                Name = "SampleModule",
                Type = VbaModuleType.ProceduralModule,
                SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from the source module!""
End Sub"
            };
            sourceDoc.VbaProject.Modules.Add(sourceModule);

            // Save the source document in macro-enabled format.
            sourceDoc.Save(sourcePath, SaveFormat.Docm);

            // ---------- Create the destination document ----------
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
            // Load the source document (optional if still in memory) to obtain the module.
            Document loadedSource = new Document(sourcePath);
            VbaModule moduleToCopy = loadedSource.VbaProject.Modules["SampleModule"];

            // Guard against null module.
            if (moduleToCopy != null)
            {
                // Clone the module and add it to the destination project's modules collection.
                VbaModule copiedModule = moduleToCopy.Clone();
                destDoc.VbaProject.Modules.Add(copiedModule);
            }

            // Save the destination document in macro-enabled format.
            destDoc.Save(destPath, SaveFormat.Docm);
        }
    }
}
