using System;
using Aspose.Words;
using Aspose.Words.Vba;

namespace AsposeWordsVbaExample
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

            // VBA code that uses late binding to a custom COM library.
            string vbaCode = @"
Sub CallCustomCom()
    Dim comObj As Object
    Set comObj = CreateObject(""MyCom.Library"")
    ' Call a method on the COM object.
    comObj.DoSomething
End Sub";

            // Create a new VBA module and set its name and source code.
            VbaModule vbaModule = new VbaModule();
            vbaModule.Name = "CustomComModule";
            vbaModule.SourceCode = vbaCode;

            // Add the module to the VBA project.
            doc.VbaProject.Modules.Add(vbaModule);

            // Save the document in macro‑enabled format.
            const string outputPath = "OutputDocument.docm";
            doc.Save(outputPath, SaveFormat.Docm);

            // Reload the document to verify the module was added.
            Document loadedDoc = new Document(outputPath);
            VbaProject loadedProject = loadedDoc.VbaProject;
            bool moduleFound = false;

            if (loadedProject != null)
            {
                foreach (VbaModule module in loadedProject.Modules)
                {
                    if (module.Name == "CustomComModule")
                    {
                        moduleFound = true;
                        // Simple validation: check that the source contains the expected method name.
                        string source = module.SourceCode ?? string.Empty;
                        if (source.Contains("CallCustomCom"))
                        {
                            // Validation succeeded.
                        }
                        break;
                    }
                }
            }

            // Output result (no interactive prompts).
            Console.WriteLine(moduleFound ? "VBA module added successfully." : "Failed to add VBA module.");
        }
    }
}
