using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

class VbaModuleCopyExample
{
    static void Main()
    {
        // Create temporary directory for the example files.
        string tempDir = Path.Combine(Path.GetTempPath(), "VbaModuleCopyExample");
        Directory.CreateDirectory(tempDir);

        // Paths to the source, destination, and result documents.
        string sourcePath = Path.Combine(tempDir, "Source.docm");
        string destinationPath = Path.Combine(tempDir, "Destination.docm");
        string resultPath = Path.Combine(tempDir, "Result.docm");

        // -----------------------------------------------------------------
        // Create a source document with a VBA project and a single module.
        // -----------------------------------------------------------------
        Document srcDoc = new Document();
        srcDoc.VbaProject = new VbaProject();

        // Create a simple VBA module.
        string moduleCode = "Sub Test()\n    MsgBox \"Hello from Source\"\nEnd Sub";
        VbaModule srcModule = new VbaModule();
        srcModule.Name = "Module1";
        srcModule.SourceCode = moduleCode;
        srcDoc.VbaProject.Modules.Add(srcModule);

        // Save the source document as a macro-enabled file.
        srcDoc.Save(sourcePath, SaveFormat.Docm);

        // -----------------------------------------------------------------
        // Create a destination document (initially without a VBA project).
        // -----------------------------------------------------------------
        Document dstDoc = new Document();
        dstDoc.Save(destinationPath, SaveFormat.Docm);

        // Reload the documents to ensure they are read from disk.
        srcDoc = new Document(sourcePath);
        dstDoc = new Document(destinationPath);

        // Ensure the destination document has a VBA project.
        if (dstDoc.VbaProject == null)
            dstDoc.VbaProject = new VbaProject();

        // Name of the VBA module to copy.
        const string moduleName = "Module1";

        // Retrieve the module from the source document.
        VbaModule moduleToCopy = srcDoc.VbaProject.Modules[moduleName];
        if (moduleToCopy != null)
        {
            // Clone the source module.
            VbaModule clonedModule = (VbaModule)moduleToCopy.Clone();

            // If a module with the same name already exists in the destination, remove it.
            VbaModule existingModule = dstDoc.VbaProject.Modules[moduleName];
            if (existingModule != null)
                dstDoc.VbaProject.Modules.Remove(existingModule);

            // Add the cloned module to the destination document's VBA project.
            dstDoc.VbaProject.Modules.Add(clonedModule);
        }

        // Save the modified destination document.
        dstDoc.Save(resultPath, SaveFormat.Docm);

        Console.WriteLine($"Source document:  {sourcePath}");
        Console.WriteLine($"Destination document (original): {destinationPath}");
        Console.WriteLine($"Result document with copied VBA module: {resultPath}");
    }
}
