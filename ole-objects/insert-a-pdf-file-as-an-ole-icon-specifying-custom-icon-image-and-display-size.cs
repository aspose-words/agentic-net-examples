using System;
using System.IO;
using System.Runtime.InteropServices;

public class Program
{
    public static void Main()
    {
        // Create temporary folder
        string tempPath = Path.Combine(Path.GetTempPath(), "OlePdfExample_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempPath);

        // Create a minimal PDF file
        string pdfPath = Path.Combine(tempPath, "sample.pdf");
        string pdfContent = @"%PDF-1.4
%âãÏÓ
1 0 obj
<< /Type /Catalog /Pages 2 0 R >>
endobj
2 0 obj
<< /Type /Pages /Kids [3 0 R] /Count 1 >>
endobj
3 0 obj
<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] >>
endobj
xref
0 4
0000000000 65535 f 
0000000010 00000 n 
0000000060 00000 n 
0000000117 00000 n 
trailer
<< /Root 1 0 R /Size 4 >>
startxref
176
%%EOF";
        File.WriteAllText(pdfPath, pdfContent);

        // Path to a custom icon (using a system icon as example)
        string iconPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), @"shell32.dll");

        // Try to get the Word COM type
        Type wordType = Type.GetTypeFromProgID("Word.Application");
        if (wordType == null)
        {
            Console.WriteLine("Microsoft Word is not installed or not registered for COM interop on this machine.");
            Console.WriteLine($"PDF created at: {pdfPath}");
            Console.WriteLine($"Temporary folder: {tempPath}");
            return;
        }

        // Start Word via COM
        dynamic wordApp = Activator.CreateInstance(wordType);
        wordApp.Visible = false;

        // Add a new document
        dynamic doc = wordApp.Documents.Add();

        // Insert the PDF as an OLE object displayed as an icon
        dynamic inlineShape = doc.InlineShapes.AddOLEObject(
            ClassType: "",
            FileName: pdfPath,
            LinkToFile: false,
            DisplayAsIcon: true,
            IconFileName: iconPath,
            IconIndex: 0,
            IconLabel: "Sample PDF"
        );

        // Set custom display size (points)
        inlineShape.Width = 200;   // approx 2.78 inches
        inlineShape.Height = 200;  // approx 2.78 inches

        // Save the document
        string docPath = Path.Combine(tempPath, "Result.docx");
        doc.SaveAs2(docPath);
        doc.Close();

        // Quit Word
        wordApp.Quit();

        // Release COM objects
        Marshal.FinalReleaseComObject(inlineShape);
        Marshal.FinalReleaseComObject(doc);
        Marshal.FinalReleaseComObject(wordApp);

        Console.WriteLine($"Word document created at: {docPath}");
    }
}
