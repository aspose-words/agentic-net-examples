using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Create a temporary VBA module file with sample code
        string vbaFilePath = "sample.bas";
        string vbaCode = "Option Explicit\r\n\r\nSub HelloWorld()\r\n    MsgBox \"Hello from VBA!\"\r\nEnd Sub";
        File.WriteAllText(vbaFilePath, vbaCode);

        // Create a new blank macro‑enabled document
        Document doc = new Document();

        // Ensure the document has a VBA project
        if (doc.VbaProject == null)
        {
            doc.VbaProject = new VbaProject();
        }

        // Read the VBA code from the external .bas file
        string importedCode = File.ReadAllText(vbaFilePath);

        // Create a new VBA module and set its properties
        VbaModule importedModule = new VbaModule
        {
            Name = "MyMacroModule",   // Desired module name
            SourceCode = importedCode // VBA source code
            // The default module type is Standard, so no need to set Type explicitly
        };

        // Add the module to the VBA project's module collection
        doc.VbaProject.Modules.Add(importedModule);

        // Save the document in a macro‑enabled format
        string outputPath = "output.docm";
        doc.Save(outputPath, SaveFormat.Docm);
    }
}
