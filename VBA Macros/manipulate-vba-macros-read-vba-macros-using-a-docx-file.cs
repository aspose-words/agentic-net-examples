using System;
using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Path to the DOCX file that may contain VBA macros.
        string filePath = "Input.docx";

        // Quick detection of macros without loading the full document.
        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
        Console.WriteLine($"Has macros (quick check): {formatInfo.HasMacros}");

        // Load the document.
        Document doc = new Document(filePath);

        // Verify that the document actually contains macros.
        if (!doc.HasMacros)
        {
            Console.WriteLine("Document does not contain any VBA macros.");
            return;
        }

        // Access the VBA project.
        VbaProject vbaProject = doc.VbaProject;
        Console.WriteLine($"VBA Project Name: {vbaProject.Name}");
        Console.WriteLine($"Modules count: {vbaProject.Modules.Count}");

        // Iterate through each VBA module and display its source code.
        foreach (VbaModule module in vbaProject.Modules)
        {
            Console.WriteLine($"--- Module: {module.Name} ---");
            Console.WriteLine(module.SourceCode);
            Console.WriteLine();
        }
    }
}
