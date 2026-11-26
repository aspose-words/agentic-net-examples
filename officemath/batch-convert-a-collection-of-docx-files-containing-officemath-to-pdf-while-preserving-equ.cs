using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class BatchOfficeMathToPdf
{
    public static void Main()
    {
        // Define input and output folders.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputPdfs");

        // Ensure folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample DOCX files containing OfficeMath equations.
        SeedSampleDocuments(inputFolder, 3);

        // Process each DOCX file: load, verify top‑level OfficeMath, and save as PDF.
        string[] docxFiles = Directory.GetFiles(inputFolder, "*.docx");
        foreach (string docxPath in docxFiles)
        {
            Document doc = new Document(docxPath);

            // Verify that the document contains at least one top‑level OfficeMath node.
            NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
            bool hasTopLevelEquation = false;
            foreach (OfficeMath om in mathNodes)
            {
                if (om.MathObjectType == MathObjectType.OMathPara)
                {
                    hasTopLevelEquation = true;
                    break;
                }
            }

            if (!hasTopLevelEquation)
                throw new InvalidOperationException($"Document '{Path.GetFileName(docxPath)}' does not contain a top‑level OfficeMath equation.");

            // Save as PDF preserving equation fidelity.
            string pdfPath = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(docxPath) + ".pdf");
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Validate that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new FileNotFoundException($"Failed to create PDF for '{docxPath}'.");
        }

        Console.WriteLine($"Batch conversion completed. PDFs are located in: {outputFolder}");
    }

    // Creates a set of sample DOCX files each containing a simple OfficeMath equation.
    private static void SeedSampleDocuments(string folderPath, int count)
    {
        for (int i = 1; i <= count; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Heading.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln($"Sample Document {i}");

            // Intro paragraph.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("The following is a simple equation:");

            // Insert an EQ field as a bootstrap for real OfficeMath.
            Field field = builder.InsertField(FieldType.FieldEquation, true);
            FieldEQ fieldEq = field as FieldEQ;
            if (fieldEq == null)
                throw new InvalidOperationException("Failed to create FieldEQ.");

            // Write a simple EQ argument (fraction 1/2) into the field separator.
            builder.MoveTo(fieldEq.Separator);
            builder.Write(@"\f(1,2)");

            // Update the field so that Aspose.Words can generate the OfficeMath node.
            fieldEq.Update();

            // Convert the field to a real OfficeMath node.
            OfficeMath officeMath = fieldEq.AsOfficeMath();
            if (officeMath != null)
            {
                // Insert the OfficeMath node before the field start and remove the original field.
                fieldEq.Start.ParentNode.InsertBefore(officeMath, fieldEq.Start);
                fieldEq.Remove();
            }
            else
            {
                throw new InvalidOperationException("EQ field could not be converted to OfficeMath.");
            }

            // Concluding paragraph.
            builder.Writeln("End of equation.");

            // Save the sample DOCX file.
            string filePath = Path.Combine(folderPath, $"Sample{i}.docx");
            doc.Save(filePath, SaveFormat.Docx);

            // Validate that the file was created.
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"Failed to create sample document '{filePath}'.");
        }
    }
}
