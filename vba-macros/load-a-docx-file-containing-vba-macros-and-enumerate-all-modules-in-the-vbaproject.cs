using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Path for the temporary macro-enabled document.
        string filePath = "Sample.docm";

        // Create a new blank document.
        Document doc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject project = new VbaProject();
        project.Name = "SampleProject";
        doc.VbaProject = project;

        // Create the first VBA module.
        VbaModule module1 = new VbaModule();
        module1.Name = "Module1";
        module1.Type = VbaModuleType.ProceduralModule;
        module1.SourceCode = "Sub Hello()\n    MsgBox \"Hello World\"\nEnd Sub";

        // Add the module to the VBA project.
        doc.VbaProject.Modules.Add(module1);

        // Create a second VBA module.
        VbaModule module2 = new VbaModule();
        module2.Name = "Module2";
        module2.Type = VbaModuleType.ProceduralModule;
        module2.SourceCode = "Function AddNumbers(a As Integer, b As Integer) As Integer\n    AddNumbers = a + b\nEnd Function";

        // Add the second module.
        doc.VbaProject.Modules.Add(module2);

        // Save the document in macro-enabled format.
        doc.Save(filePath);

        // Load the saved document.
        Document loadedDoc = new Document(filePath);

        // Access the VBA project.
        VbaProject loadedProject = loadedDoc.VbaProject;

        // If the document contains a VBA project, enumerate its modules.
        if (loadedProject != null)
        {
            Console.WriteLine($"VBA Project Name: {loadedProject.Name}");
            Console.WriteLine($"Modules Count: {loadedProject.Modules.Count}");

            foreach (VbaModule vbaModule in loadedProject.Modules)
            {
                // Guard against null source code.
                string source = vbaModule.SourceCode ?? string.Empty;

                Console.WriteLine($"--- Module: {vbaModule.Name} ---");
                Console.WriteLine(source);
                Console.WriteLine();
            }
        }
        else
        {
            Console.WriteLine("The document does not contain a VBA project.");
        }

        // Clean up the temporary file.
        if (File.Exists(filePath))
        {
            File.Delete(filePath);
        }
    }
}
