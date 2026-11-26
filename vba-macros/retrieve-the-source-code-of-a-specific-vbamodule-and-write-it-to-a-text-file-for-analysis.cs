using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

namespace VbaModuleExtractor
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Ensure the document has a VBA project.
            if (doc.VbaProject == null)
                doc.VbaProject = new VbaProject();

            // Add a VBA module with sample code.
            string moduleName = "SampleModule";
            string moduleCode = @"Sub HelloWorld()
    MsgBox ""Hello, World!""
End Sub";

            // VbaModule does not have a constructor that accepts name and code.
            // Create an empty module and set its properties.
            VbaModule vbaModule = new VbaModule
            {
                Name = moduleName,
                SourceCode = moduleCode
            };

            doc.VbaProject.Modules.Add(vbaModule);

            // Save the document in macro‑enabled format.
            string docPath = "sample.docm";
            doc.Save(docPath, SaveFormat.Docm);

            // Load the saved document.
            Document loadedDoc = new Document(docPath);

            // Retrieve the specified VBA module.
            VbaModule targetModule = null;
            if (loadedDoc.VbaProject != null && loadedDoc.VbaProject.Modules != null)
                targetModule = loadedDoc.VbaProject.Modules[moduleName];

            // Guard against null source code.
            string sourceCode = targetModule?.SourceCode ?? string.Empty;

            // Write the source code to a text file.
            string outputPath = "module_source.txt";
            File.WriteAllText(outputPath, sourceCode);
        }
    }
}
