using System;
using Aspose.Words;
using Aspose.Words.Vba;

class UpdateVbaModuleSource
{
    static void Main()
    {
        // Path to the input DOCM file that contains VBA macros.
        const string inputPath = @"C:\Docs\InputDocument.docm";
        // Path where the updated document will be saved.
        const string outputPath = @"C:\Docs\OutputDocument.docm";

        // Load the document (creates a Document instance).
        Document doc = new Document(inputPath);

        // Ensure the document actually has a VBA project.
        if (doc.VbaProject == null)
        {
            Console.WriteLine("The document does not contain a VBA project.");
            return;
        }

        // Retrieve the first VBA module in the project (you can also use a name or index).
        VbaModuleCollection modules = doc.VbaProject.Modules;
        if (modules.Count == 0)
        {
            Console.WriteLine("No VBA modules found in the project.");
            return;
        }

        VbaModule module = modules[0]; // or modules["ModuleName"]

        // Get the existing source code.
        string originalSource = module.SourceCode ?? string.Empty;

        // Prepare the error‑handling statement.
        const string errorHandling = "On Error Resume Next";

        // If the source does not already start with the error handling line, prepend it.
        if (!originalSource.TrimStart().StartsWith(errorHandling, StringComparison.OrdinalIgnoreCase))
        {
            // Use Windows line endings for consistency with VBA.
            string updatedSource = errorHandling + "\r\n" + originalSource;
            module.SourceCode = updatedSource;
            Console.WriteLine($"Updated source code of module '{module.Name}'.");
        }
        else
        {
            Console.WriteLine($"Module '{module.Name}' already contains error handling.");
        }

        // Save the modified document (preserves the .docm macro format).
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to '{outputPath}'.");
    }
}
