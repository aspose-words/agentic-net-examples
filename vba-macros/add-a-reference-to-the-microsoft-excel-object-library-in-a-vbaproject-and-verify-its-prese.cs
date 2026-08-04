using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Define output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Path for the macro-enabled document.
        string docPath = Path.Combine(outputDir, "ExcelReference.docm");

        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has a VBA project.
        VbaProject vbaProject = new VbaProject();
        vbaProject.Name = "Aspose.ExcelProject";
        doc.VbaProject = vbaProject;

        // Add a simple VBA module (the content is not important for this demo).
        VbaModule module = new VbaModule
        {
            Name = "Module1",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = "Sub Dummy()\n    MsgBox \"Hello\"\nEnd Sub"
        };
        doc.VbaProject.Modules.Add(module);

        // Save the document in macro-enabled format.
        doc.Save(docPath);

        // Reload the document to work with the saved VBA project.
        Document loadedDoc = new Document(docPath);

        // Access the VBA project.
        VbaProject loadedProject = loadedDoc.VbaProject;

        // Verify the presence of a reference to the Microsoft Excel Object Library.
        // The reference type is usually Registered and its LibId contains "EXCEL".
        bool hasExcelReference = false;
        foreach (VbaReference reference in loadedProject.References)
        {
            // Guard against null LibId.
            string libId = reference.LibId ?? string.Empty;
            if (libId.IndexOf("EXCEL", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                hasExcelReference = true;
                break;
            }
        }

        // Output verification result.
        Console.WriteLine(hasExcelReference
            ? "Microsoft Excel Object Library reference is present."
            : "Microsoft Excel Object Library reference is NOT present.");
    }
}
