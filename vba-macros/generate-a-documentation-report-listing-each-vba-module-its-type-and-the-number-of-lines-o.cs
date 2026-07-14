using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Define file paths.
        string docPath = Path.Combine(Directory.GetCurrentDirectory(), "MacroDocument.docm");
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "VbaModulesReport.txt");

        // -----------------------------------------------------------------
        // 1. Create a new blank document and a VBA project with sample modules.
        // -----------------------------------------------------------------
        Document doc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject vbaProject = new VbaProject
        {
            Name = "SampleProject"
        };
        doc.VbaProject = vbaProject;

        // Helper to add a module.
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
            "ClassHelper",
            VbaModuleType.ClassModule,
            @"Option Explicit

Public Sub ShowMessage()
    MsgBox ""Message from ClassHelper!""
End Sub"
        );

        // Save the macro-enabled document.
        doc.Save(docPath, SaveFormat.Docm);

        // -----------------------------------------------------------------
        // 2. Load the document and generate a report of VBA modules.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);

        // Ensure the document actually contains macros.
        if (!loadedDoc.HasMacros || loadedDoc.VbaProject == null)
        {
            Console.WriteLine("The document does not contain any VBA project.");
            return;
        }

        VbaModuleCollection modules = loadedDoc.VbaProject.Modules;

        using (StreamWriter writer = new StreamWriter(reportPath))
        {
            writer.WriteLine("VBA Module Report");
            writer.WriteLine("-----------------");
            writer.WriteLine($"Document: {Path.GetFileName(docPath)}");
            writer.WriteLine();

            foreach (VbaModule module in modules)
            {
                // Guard against null source code.
                string source = module.SourceCode ?? string.Empty;

                // Count lines (including empty lines) by splitting on newline characters.
                int lineCount = source.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Length;

                string line = $"Name: {module.Name}, Type: {module.Type}, Lines of Code: {lineCount}";
                Console.WriteLine(line);
                writer.WriteLine(line);
            }
        }

        // Output location of the generated report.
        Console.WriteLine($"Report saved to: {reportPath}");
    }
}
