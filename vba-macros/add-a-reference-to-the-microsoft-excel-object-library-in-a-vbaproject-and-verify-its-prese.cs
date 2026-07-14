using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Define file paths for the temporary macro-enabled document.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string docPath = Path.Combine(outputDir, "ExcelReferenceDemo.docm");

        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has a VBA project.
        VbaProject vbaProject = new VbaProject();
        vbaProject.Name = "ExcelReferenceProject";
        doc.VbaProject = vbaProject;

        // Add a simple VBA module (the actual code is not important for this demo).
        VbaModule module = new VbaModule
        {
            Name = "DemoModule",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = "Sub Dummy()\nEnd Sub"
        };
        doc.VbaProject.Modules.Add(module);

        // Save the document in a macro‑enabled format.
        doc.Save(docPath);

        // Reload the document to work with the saved VBA project.
        Document loadedDoc = new Document(docPath);

        // Verify that the document indeed contains a VBA project.
        if (!loadedDoc.HasMacros || loadedDoc.VbaProject == null)
        {
            Console.WriteLine("The document does not contain a VBA project.");
            return;
        }

        // Check the VBA references collection for the Microsoft Excel Object Library.
        bool excelReferenceFound = false;
        foreach (VbaReference reference in loadedDoc.VbaProject.References)
        {
            // The LibId string typically contains the library name; look for "EXCEL".
            if (!string.IsNullOrEmpty(reference.LibId) &&
                reference.LibId.IndexOf("EXCEL", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                excelReferenceFound = true;
                break;
            }
        }

        // Output the verification result.
        Console.WriteLine(excelReferenceFound
            ? "Microsoft Excel Object Library reference is present."
            : "Microsoft Excel Object Library reference is NOT present.");
    }
}
