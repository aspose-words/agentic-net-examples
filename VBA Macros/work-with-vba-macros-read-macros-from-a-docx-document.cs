using System;
using Aspose.Words;
using Aspose.Words.Vba;

class ReadVbaMacros
{
    static void Main()
    {
        // Load the document that may contain VBA macros.
        // Adjust the file path as needed; the document must be a macro‑enabled format (e.g., .docm).
        Document doc = new Document("Input.docm");

        // Verify that the document actually has a VBA project.
        if (!doc.HasMacros)
        {
            Console.WriteLine("The document does not contain any VBA macros.");
            return;
        }

        // Access the VBA project.
        VbaProject vbaProject = doc.VbaProject;

        // Iterate through each VBA module and output its name and source code.
        foreach (VbaModule module in vbaProject.Modules)
        {
            Console.WriteLine($"Module Name: {module.Name}");
            Console.WriteLine("Source Code:");
            Console.WriteLine(module.SourceCode);
            Console.WriteLine(new string('-', 40));
        }
    }
}
