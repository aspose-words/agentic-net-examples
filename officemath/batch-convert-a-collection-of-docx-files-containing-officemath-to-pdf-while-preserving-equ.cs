using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class OfficeMathBatchConverter
{
    private static void CreateSampleDoc(string filePath, string equation)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert an equation field.
        var field = builder.InsertField(FieldType.FieldEquation, true);
        // Write the EQ argument.
        builder.MoveTo(field.Separator);
        builder.Write(equation);

        // Convert the field to a real OfficeMath node.
        if (field is FieldEQ fieldEq)
        {
            var officeMath = fieldEq.AsOfficeMath();
            if (officeMath != null)
            {
                // Insert the OfficeMath node before the field start.
                var startNode = field.Start;
                startNode.ParentNode.InsertBefore(officeMath, startNode);
                // Remove the original field (start, separator, end).
                field.Remove();
            }
        }

        doc.Save(filePath, SaveFormat.Docx);
    }

    public static void Main()
    {
        // Prepare input and output folders.
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string outputFolder = Path.Combine(baseDir, "OutputPdfs");

        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample DOCX files with OfficeMath equations.
        CreateSampleDoc(Path.Combine(inputFolder, "Doc1.docx"), @"\f(1,2)");
        CreateSampleDoc(Path.Combine(inputFolder, "Doc2.docx"), @"\r(3,x)");
        CreateSampleDoc(Path.Combine(inputFolder, "Doc3.docx"), @"\f(5,7)");

        // Batch convert each DOCX to PDF.
        foreach (string docPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            var doc = new Document(docPath);
            string pdfPath = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(docPath) + ".pdf");
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Validate that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new Exception($"PDF conversion failed for '{docPath}'.");
        }

        // Simple verification that all PDFs exist.
        int pdfCount = Directory.GetFiles(outputFolder, "*.pdf").Length;
        if (pdfCount == 0)
            throw new Exception("No PDF files were generated.");

        // Indicate successful completion.
        Console.WriteLine($"Batch conversion completed. {pdfCount} PDF files created in '{outputFolder}'.");
    }
}
