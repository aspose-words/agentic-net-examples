using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Define folders for input DOCX files and output PDF files.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputPdfs");

        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample DOCX files that contain OfficeMath equations.
        for (int i = 1; i <= 3; i++)
        {
            string docPath = Path.Combine(inputFolder, $"Sample{i}.docx");
            CreateSampleDocumentWithOfficeMath(docPath);
        }

        // Batch convert each DOCX file to PDF while preserving equation fidelity.
        foreach (string docxFile in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the DOCX document.
            Document doc = new Document(docxFile);

            // Determine the corresponding PDF file path.
            string pdfFileName = Path.GetFileNameWithoutExtension(docxFile) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Save as PDF.
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Validate that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF: {pdfPath}");
        }

        // Optional: indicate completion.
        Console.WriteLine("Batch conversion completed successfully.");
    }

    // Creates a DOCX file at the specified path containing a simple OfficeMath equation.
    private static void CreateSampleDocumentWithOfficeMath(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an EQ field with a simple fraction equation.
        FieldEQ field = InsertFieldEQ(builder, @"\f(1,2)");

        // Convert the EQ field to a real OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();

        // Insert the OfficeMath node before the field and remove the original field.
        if (officeMath != null)
        {
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            field.Remove();
        }

        // Save the document as DOCX.
        doc.Save(filePath, SaveFormat.Docx);
    }

    // Helper that inserts an EQ field, writes the argument string, and moves the builder back to the paragraph.
    private static FieldEQ InsertFieldEQ(DocumentBuilder builder, string args)
    {
        // Insert the EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Move to the field separator and write the equation arguments.
        builder.MoveTo(field.Separator);
        builder.Write(args);

        // Return the builder to the paragraph after the field.
        builder.MoveTo(field.Start.ParentNode);
        builder.InsertParagraph();

        return field;
    }
}
