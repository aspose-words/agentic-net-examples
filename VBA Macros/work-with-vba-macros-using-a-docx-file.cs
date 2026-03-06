using System;
using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = "input.docx";

        // Path to the macro‑enabled output file.
        string outputPath = "output.docm";

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // Check whether the document already contains VBA macros.
        if (doc.HasMacros)
        {
            Console.WriteLine("Document already contains macros.");
        }
        else
        {
            Console.WriteLine("Document does not contain macros. Adding a new VBA project.");

            // Create a new blank VBA project.
            VbaProject vbaProject = new VbaProject();
            vbaProject.Name = "AsposeProject";

            // Assign the VBA project to the document.
            doc.VbaProject = vbaProject;

            // Create a new procedural module.
            VbaModule module = new VbaModule();
            module.Name = "Module1";
            module.Type = VbaModuleType.ProceduralModule;

            // Simple macro source code.
            module.SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from Aspose.Words!""
End Sub
";

            // Add the module to the VBA project.
            doc.VbaProject.Modules.Add(module);
        }

        // Save the document as a macro‑enabled file.
        doc.Save(outputPath, SaveFormat.Docm);

        // Load the saved document to demonstrate macro removal.
        Document docWithMacros = new Document(outputPath);
        Console.WriteLine($"Has macros before removal: {docWithMacros.HasMacros}");

        // Remove all macros, toolbars, and customizations.
        docWithMacros.RemoveMacros();

        // Save the cleaned document as a regular DOCX.
        string cleanedPath = "cleaned.docx";
        docWithMacros.Save(cleanedPath, SaveFormat.Docx);
        Console.WriteLine($"Has macros after removal: {docWithMacros.HasMacros}");
    }
}
