using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Define paths for the temporary macro-enabled document and the report output.
        string docPath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.docm");

        // -----------------------------------------------------------------
        // Create a new blank document and a VBA project with sample modules.
        // -----------------------------------------------------------------
        Document doc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject project = new VbaProject
        {
            Name = "SampleProject"
        };
        doc.VbaProject = project;

        // Helper to add a module with given name, type and source code.
        void AddModule(string name, VbaModuleType type, string source)
        {
            VbaModule module = new VbaModule
            {
                Name = name,
                Type = type,
                SourceCode = source
            };
            doc.VbaProject.Modules.Add(module);
        }

        // Add a procedural module.
        AddModule(
            "ModuleProcedural",
            VbaModuleType.ProceduralModule,
            @"Sub HelloWorld()
    MsgBox ""Hello, World!""
End Sub"
        );

        // Add a class module.
        AddModule(
            "ClassExample",
            VbaModuleType.ClassModule,
            @"Option Explicit

Public Sub ShowMessage()
    MsgBox ""Message from class module.""
End Sub"
        );

        // Add a document module.
        AddModule(
            "DocumentModule",
            VbaModuleType.DocumentModule,
            @"Sub AutoOpen()
    MsgBox ""Document opened.""
End Sub"
        );

        // Save the document in macro-enabled format.
        doc.Save(docPath, SaveFormat.Docm);

        // ---------------------------------------------------------------
        // Load the document back and generate a report of its VBA modules.
        // ---------------------------------------------------------------
        Document loadedDoc = new Document(docPath);

        // Ensure the document actually contains a VBA project.
        if (loadedDoc.HasMacros && loadedDoc.VbaProject != null)
        {
            VbaModuleCollection modules = loadedDoc.VbaProject.Modules;

            Console.WriteLine("VBA Modules Report:");
            Console.WriteLine("-------------------");

            foreach (VbaModule module in modules)
            {
                // Guard against null source code.
                string source = module.SourceCode ?? string.Empty;

                // Count lines of code. Handles CRLF, LF, or CR line endings.
                int lineCount = source.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None).Length;

                // Output module details.
                Console.WriteLine($"Name: {module.Name}");
                Console.WriteLine($"Type: {module.Type}");
                Console.WriteLine($"Lines of Code: {lineCount}");
                Console.WriteLine();
            }
        }
        else
        {
            Console.WriteLine("The document does not contain any VBA macros.");
        }

        // Clean up the temporary file (optional).
        // File.Delete(docPath);
    }
}
