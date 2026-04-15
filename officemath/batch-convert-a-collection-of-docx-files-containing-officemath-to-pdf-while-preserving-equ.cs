using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main(string[] args)
    {
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string outputFolder = Path.Combine(baseDir, "OutputPdfs");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample DOCX files with simple OfficeMath equations
        for (int i = 1; i <= 2; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln($"Sample document {i} with an equation:");

            // Insert an EQ field
            Field field = builder.InsertField(FieldType.FieldEquation, true);
            FieldEQ fieldEQ = field as FieldEQ;
            if (fieldEQ == null)
                throw new InvalidOperationException("Failed to create FieldEQ.");

            // Write a simple EQ argument (deterministic safe expression)
            builder.MoveTo(fieldEQ.Separator);
            builder.Write(@"\f(1,2)");

            // Update the field so that Aspose.Words can convert it to OfficeMath
            fieldEQ.Update();

            // Convert the field to a real OfficeMath node
            OfficeMath officeMath = fieldEQ.AsOfficeMath();
            if (officeMath == null)
                throw new InvalidOperationException("EQ field could not be converted to OfficeMath.");

            // Insert the OfficeMath node before the field start and remove the field
            Node fieldStart = field.Start;
            if (fieldStart == null)
                throw new InvalidOperationException("Field start node is null.");

            fieldStart.ParentNode.InsertBefore(officeMath, fieldStart);
            field.Remove();

            builder.Writeln(); // Add a blank line after the equation

            string docPath = Path.Combine(inputFolder, $"Sample{i}.docx");
            doc.Save(docPath, SaveFormat.Docx);
        }

        // Batch convert all DOCX files in the input folder to PDF
        string[] docxFiles = Directory.GetFiles(inputFolder, "*.docx");
        foreach (string docxPath in docxFiles)
        {
            Document doc = new Document(docxPath);
            string pdfFileName = Path.GetFileNameWithoutExtension(docxPath) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Validate that the PDF was created
            if (!File.Exists(pdfPath))
                throw new FileNotFoundException($"Failed to create PDF: {pdfPath}");
        }

        // Indicate successful completion
        Console.WriteLine($"Converted {docxFiles.Length} DOCX file(s) to PDF in '{outputFolder}'.");
    }
}
