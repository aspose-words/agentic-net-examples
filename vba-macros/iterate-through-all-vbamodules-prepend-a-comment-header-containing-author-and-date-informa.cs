using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Define author information.
        const string author = "John Doe";

        // Create a new blank document.
        Document doc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject vbaProject = new VbaProject();
        vbaProject.Name = "SampleProject";
        doc.VbaProject = vbaProject;

        // Add a couple of VBA modules with sample code.
        VbaModule module1 = new VbaModule();
        module1.Name = "Module1";
        module1.Type = VbaModuleType.ProceduralModule;
        module1.SourceCode = "Sub HelloWorld()\n    MsgBox \"Hello, World!\"\nEnd Sub";
        doc.VbaProject.Modules.Add(module1);

        VbaModule module2 = new VbaModule();
        module2.Name = "Module2";
        module2.Type = VbaModuleType.ProceduralModule;
        module2.SourceCode = "Function AddNumbers(a As Integer, b As Integer) As Integer\n    AddNumbers = a + b\nEnd Function";
        doc.VbaProject.Modules.Add(module2);

        // Save the initial document (optional, shows the original state).
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "Original.docm");
        doc.Save(originalPath, SaveFormat.Docm);

        // Iterate through all VBA modules and prepend a comment header.
        foreach (VbaModule vbaModule in doc.VbaProject.Modules)
        {
            // Guard against null source code.
            string existingCode = vbaModule.SourceCode ?? string.Empty;

            // Build the header comment.
            string header = $"' Author: {author}, Date: {DateTime.Now:yyyy-MM-dd}\r\n";

            // Prepend the header to the existing source code.
            vbaModule.SourceCode = header + existingCode;
        }

        // Save the modified document.
        string modifiedPath = Path.Combine(Directory.GetCurrentDirectory(), "Modified.docm");
        doc.Save(modifiedPath, SaveFormat.Docm);
    }
}
