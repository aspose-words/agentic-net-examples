using System;
using Aspose.Words;
using Aspose.Words.Vba;

class ExploreVbaMacros
{
    static void Main()
    {
        // Load a macro-enabled document (DOCM). The constructor automatically detects the format.
        Document doc = new Document("VBA project.docm");

        // Check if the document actually contains VBA macros.
        if (!doc.HasMacros)
        {
            Console.WriteLine("The document does not contain any VBA macros.");
            return;
        }

        // Access the VBA project associated with the document.
        VbaProject vbaProject = doc.VbaProject;

        // Output basic information about the VBA project.
        Console.WriteLine(vbaProject.IsSigned
            ? $"Project name: {vbaProject.Name} (signed); Code page: {vbaProject.CodePage}"
            : $"Project name: {vbaProject.Name} (not signed); Code page: {vbaProject.CodePage}");

        // Get the collection of VBA modules.
        VbaModuleCollection modules = vbaProject.Modules;
        Console.WriteLine($"Modules count: {modules.Count}");

        // Iterate through each module and display its name and source code.
        foreach (VbaModule module in modules)
        {
            Console.WriteLine($"--- Module: {module.Name} ---");
            Console.WriteLine(module.SourceCode);
            Console.WriteLine();
        }

        // Example: modify the source code of the first module (if any).
        if (modules.Count > 0)
        {
            VbaModule firstModule = modules[0];
            firstModule.SourceCode = "' Updated VBA code\r\nSub HelloWorld()\r\n    MsgBox \"Hello from Aspose.Words!\"\r\nEnd Sub";
            Console.WriteLine($"Updated source code of module '{firstModule.Name}'.");
        }

        // Save the document with the modified macros to a new file.
        doc.Save("VBA project - Modified.docm");
        Console.WriteLine("Document saved with updated VBA macros.");
    }
}
