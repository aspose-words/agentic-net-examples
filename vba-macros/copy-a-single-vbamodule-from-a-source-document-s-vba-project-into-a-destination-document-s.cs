using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Vba;

namespace VbaModuleCopyExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a source macro‑enabled document.
            Document sourceDoc = new Document();

            // Ensure the source document has a VBA project.
            sourceDoc.VbaProject = new VbaProject();

            // Define a simple VBA module.
            string moduleName = "TestModule";
            string moduleCode = "Sub Hello()\n    MsgBox \"Hello from source\"\nEnd Sub";

            // Create the module, set its name and source code, then add it to the source document's VBA project.
            VbaModule sourceModule = new VbaModule();
            sourceModule.Name = moduleName;
            sourceModule.SourceCode = moduleCode;
            sourceDoc.VbaProject.Modules.Add(sourceModule);

            // Save the source document in a macro‑enabled format.
            sourceDoc.Save("source.docm", SaveFormat.Docm);

            // Create a destination macro‑enabled document.
            Document destDoc = new Document();

            // Ensure the destination document has a VBA project.
            destDoc.VbaProject = new VbaProject();

            // Retrieve the source module's code (guard against null).
            string sourceCode = sourceModule?.SourceCode ?? string.Empty;

            // Create a new module in the destination document using the source code.
            VbaModule destModule = new VbaModule();
            destModule.Name = sourceModule.Name;
            destModule.SourceCode = sourceCode;
            destDoc.VbaProject.Modules.Add(destModule);

            // Save the destination document in a macro‑enabled format.
            destDoc.Save("dest.docm", SaveFormat.Docm);

            // Validation: check that the module exists and its code matches.
            VbaModule copiedModule = destDoc.VbaProject.Modules[moduleName];
            bool moduleExists = copiedModule != null;
            bool codeMatches = copiedModule?.SourceCode == sourceCode;

            Console.WriteLine($"Module copied: {moduleExists && codeMatches}");
        }
    }
}
