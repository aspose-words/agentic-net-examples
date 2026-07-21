using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input DOCX files and output PDF files.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputPdfs");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Number of sample documents to create.
        const int sampleCount = 3;

        // Create sample DOCX files that contain OfficeMath equations.
        for (int i = 1; i <= sampleCount; i++)
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a simple paragraph before the equation.
            builder.Writeln($"Sample document {i} with an equation:");

            // Insert an EQ field (fraction 1/2) using the deterministic bootstrap workflow.
            FieldEQ eqField = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);
            // Move to the field separator and write the EQ argument.
            builder.MoveTo(eqField.Separator);
            builder.Write(@"\f(1,2)");
            // Return to the field start's parent node (the paragraph) and add a new paragraph after the equation.
            builder.MoveTo(eqField.Start.ParentNode);
            builder.InsertParagraph();

            // Convert the EQ field to a real OfficeMath object.
            OfficeMath officeMath = eqField.AsOfficeMath();
            if (officeMath != null)
            {
                // Insert the OfficeMath node before the field start and remove the original field.
                eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);
                eqField.Remove();
            }

            // Save the document as DOCX in the input folder.
            string docPath = Path.Combine(inputFolder, $"Sample{i}.docx");
            doc.Save(docPath, SaveFormat.Docx);
        }

        // Batch convert each DOCX file in the input folder to PDF, preserving equation fidelity.
        foreach (string docxPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the DOCX document.
            Document doc = new Document(docxPath);

            // Determine the output PDF path.
            string pdfPath = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(docxPath) + ".pdf");

            // Save the document as PDF.
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Validate that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF: {pdfPath}");
        }

        // Optional: indicate completion (no interactive prompts).
        Console.WriteLine("Batch conversion completed successfully.");
    }
}
