using System;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject vbaProject = new VbaProject();
        vbaProject.Name = "SampleProject";
        doc.VbaProject = vbaProject;

        // Add a sample procedural module.
        VbaModule module1 = new VbaModule();
        module1.Name = "Module1";
        module1.Type = VbaModuleType.ProceduralModule;
        module1.SourceCode = "Sub HelloWorld()\r\n    MsgBox \"Hello, World!\"\r\nEnd Sub";
        doc.VbaProject.Modules.Add(module1);

        // Add another sample procedural module.
        VbaModule module2 = new VbaModule();
        module2.Name = "Module2";
        module2.Type = VbaModuleType.ProceduralModule;
        module2.SourceCode = "Sub AddNumbers()\r\n    Dim a As Integer, b As Integer\r\n    a = 5\r\n    b = 10\r\n    MsgBox a + b\r\nEnd Sub";
        doc.VbaProject.Modules.Add(module2);

        // Define author and date header.
        string author = "John Doe";
        string date = DateTime.Now.ToString("yyyy-MM-dd");
        string header = $"' Author: {author}\r\n' Date: {date}\r\n";

        // Iterate through all VBA modules and prepend the header.
        foreach (VbaModule module in doc.VbaProject.Modules)
        {
            // Guard against null source code.
            string source = module.SourceCode ?? string.Empty;
            module.SourceCode = header + source;
        }

        // Save the document in a macro-enabled format.
        doc.Save("Output.docm");
    }
}
