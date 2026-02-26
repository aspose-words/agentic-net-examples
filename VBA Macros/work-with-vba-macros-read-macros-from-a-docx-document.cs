using System;
using Aspose.Words;
using Aspose.Words.Vba;

class ReadMacrosFromDocx
{
    static void Main()
    {
        // Path to the DOCX (or DOCM) file that may contain VBA macros.
        string filePath = @"C:\Docs\MacroDocument.docx";

        // Detect the file format and check for macros without fully loading the document.
        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
        Console.WriteLine($"Has macros (detected without loading): {formatInfo.HasMacros}");

        // Load the document only if we need to inspect the macro code.
        Document doc = new Document(filePath);

        // Verify that the loaded document reports macros.
        Console.WriteLine($"Has macros (after loading): {doc.HasMacros}");

        // If there are no macros, exit early.
        if (!doc.HasMacros || doc.VbaProject == null)
        {
            Console.WriteLine("No VBA project found in the document.");
            return;
        }

        // Access the VBA project.
        VbaProject vbaProject = doc.VbaProject;
        Console.WriteLine($"VBA Project Name: {vbaProject.Name}");
        Console.WriteLine($"Is Signed: {vbaProject.IsSigned}");
        Console.WriteLine($"Modules Count: {vbaProject.Modules.Count}");

        // Iterate through each VBA module and output its name and source code.
        foreach (VbaModule module in vbaProject.Modules)
        {
            Console.WriteLine($"--- Module: {module.Name} (Type: {module.Type}) ---");
            Console.WriteLine(module.SourceCode);
            Console.WriteLine();
        }
    }
}
