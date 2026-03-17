using System;
using Aspose.Words;
using Aspose.Words.Vba;

class VbaProjectValidator
{
    /// <summary>
    /// Loads a document, verifies that its VBA project contains at least one module,
    /// performs a simple modification, and saves the result.
    /// </summary>
    /// <param name="inputPath">Path to the source document (e.g., *.docm).</param>
    /// <param name="outputPath">Path where the modified document will be saved.</param>
    public static void ValidateAndModify(string inputPath, string outputPath)
    {
        // Load the document from disk.
        Document doc = new Document(inputPath);

        // Ensure the document actually has a VBA project.
        if (doc.VbaProject == null)
            throw new InvalidOperationException("The document does not contain a VBA project.");

        // Retrieve the collection of VBA modules.
        VbaModuleCollection modules = doc.VbaProject.Modules;

        // Validate that there is at least one module in the project.
        if (modules == null || modules.Count == 0)
            throw new InvalidOperationException("The VBA project contains no modules.");

        // Example modification: prepend a comment to the source code of the first module.
        VbaModule firstModule = modules[0];
        firstModule.SourceCode = "' Modified by Aspose.Words\n" + firstModule.SourceCode;

        // Save the modified document.
        doc.Save(outputPath);
    }
}

class Program
{
    static void Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: <program> <inputPath> <outputPath>");
            return;
        }

        try
        {
            VbaProjectValidator.ValidateAndModify(args[0], args[1]);
            Console.WriteLine("Document processed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
