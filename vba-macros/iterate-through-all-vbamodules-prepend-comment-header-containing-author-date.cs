using System;
using Aspose.Words;
using Aspose.Words.Vba;

class VbaHeaderAppender
{
    static void Main()
    {
        // Path to the input macro‑enabled document.
        const string inputPath = @"C:\Docs\Input.docm";
        // Path where the modified document will be saved.
        const string outputPath = @"C:\Docs\Output.docm";

        // Load the document. Aspose.Words automatically detects the format from the extension.
        Document doc = new Document(inputPath);

        // Ensure the document actually contains VBA macros before accessing the project.
        if (!doc.HasMacros)
        {
            Console.WriteLine("The document does not contain any VBA macros.");
            return;
        }

        // Author information to prepend.
        const string author = "John Doe";

        // Iterate through each VBA module in the project.
        foreach (VbaModule module in doc.VbaProject.Modules)
        {
            // Preserve the existing source code.
            string originalCode = module.SourceCode ?? string.Empty;

            // Build the comment header. VBA comments start with a single quote (').
            string header = $"' Author: {author}\r\n' Date: {DateTime.Now:yyyy-MM-dd}\r\n";

            // Prepend the header to the original source code.
            module.SourceCode = header + originalCode;
        }

        // Save the modified document. The .docm extension keeps the file macro‑enabled.
        doc.Save(outputPath);
        Console.WriteLine($"Document saved with updated VBA headers to: {outputPath}");
    }
}
