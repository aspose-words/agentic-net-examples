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

        // 1. Create a blank document.
        Document doc = new Document();

        // 2. Create a new VBA project and assign it to the document.
        VbaProject vbaProject = new VbaProject
        {
            Name = "SampleProject"
        };
        doc.VbaProject = vbaProject;

        // 3. Add sample VBA modules.
        AddVbaModule(doc, "ModuleOne", VbaModuleType.ProceduralModule,
            "Sub HelloWorld()\r\n    MsgBox \"Hello, World!\"\r\nEnd Sub");

        AddVbaModule(doc, "ClassModuleExample", VbaModuleType.ClassModule,
            "Option Explicit\r\n\r\nPublic Sub Greet()\r\n    MsgBox \"Greetings from class module.\"\r\nEnd Sub");

        // 4. Save the document in macro-enabled format.
        doc.Save(docPath, SaveFormat.Docm);

        // 5. Reload the document to demonstrate reading.
        Document loadedDoc = new Document(docPath);

        // 6. Generate report of VBA modules.
        Console.WriteLine("VBA Modules Report:");
        Console.WriteLine("--------------------");

        if (loadedDoc.HasMacros && loadedDoc.VbaProject != null)
        {
            VbaModuleCollection modules = loadedDoc.VbaProject.Modules;
            foreach (VbaModule module in modules)
            {
                // Guard against null source code.
                string source = module.SourceCode ?? string.Empty;
                // Count lines (ignore empty lines).
                int lineCount = source.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Length;

                Console.WriteLine($"Name: {module.Name}");
                Console.WriteLine($"Type: {module.Type}");
                Console.WriteLine($"Lines of Code: {lineCount}");
                Console.WriteLine();
            }
        }
        else
        {
            Console.WriteLine("No VBA macros found in the document.");
        }

        // Optional: clean up the generated file.
        if (File.Exists(docPath))
        {
            File.Delete(docPath);
        }
    }

    private static void AddVbaModule(Document doc, string name, VbaModuleType type, string sourceCode)
    {
        VbaModule module = new VbaModule
        {
            Name = name,
            Type = type,
            SourceCode = sourceCode
        };
        doc.VbaProject.Modules.Add(module);
    }
}
