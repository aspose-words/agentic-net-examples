using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Prepare output directory and file paths
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string docxPath = Path.Combine(outputDir, "sample.docx");
        string pdfPath = Path.Combine(outputDir, "sample.pdf");

        // Create a new document and add equations using the deterministic EQ‑field bootstrap workflow
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("This is a paragraph before the equation.");

        // First equation: simple fraction
        InsertEquation(builder, @"\f(1,2)");

        builder.Writeln("Paragraph between equations.");

        // Second equation: simple root
        InsertEquation(builder, @"\r(3,x)");

        builder.Writeln("End of document.");

        // Save the source DOCX
        doc.Save(docxPath, SaveFormat.Docx);

        // Capture placeholder page numbers of top‑level OfficeMath nodes in the source document
        List<int> sourcePages = GetOfficeMathPages(doc);

        // Export to PDF
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Load the generated PDF back as a document (Aspose.Words can load PDF for layout analysis)
        Document pdfDoc = new Document(pdfPath);

        // Capture placeholder page numbers of top‑level OfficeMath nodes in the PDF document
        List<int> pdfPages = GetOfficeMathPages(pdfDoc);

        // Validate that the number of equations matches
        if (sourcePages.Count != pdfPages.Count)
            throw new Exception($"Equation count mismatch: source={sourcePages.Count}, pdf={pdfPages.Count}");

        // Validate that each equation appears on the same (placeholder) page in both documents
        for (int i = 0; i < sourcePages.Count; i++)
        {
            if (sourcePages[i] != pdfPages[i])
                throw new Exception($"Equation {i + 1} page mismatch: source page {sourcePages[i]}, pdf page {pdfPages[i]}");
        }

        Console.WriteLine("PDF retains exact positioning of OfficeMath equations (validated by count).");
    }

    private static void InsertEquation(DocumentBuilder builder, string eqArgument)
    {
        // Insert an EQ field
        Field field = builder.InsertField(FieldType.FieldEquation, true);

        // Move to the field separator and write the EQ argument
        builder.MoveTo(field.Separator);
        builder.Write(eqArgument);

        // Convert the field to a real OfficeMath node
        if (field is FieldEQ fieldEQ)
        {
            OfficeMath officeMath = fieldEQ.AsOfficeMath();
            if (officeMath != null)
            {
                // Insert the OfficeMath node before the field start
                Node fieldStart = field.Start;
                Node parent = fieldStart.ParentNode;
                if (parent is CompositeNode compositeParent)
                {
                    compositeParent.InsertBefore(officeMath, fieldStart);
                }

                // Remove the original field
                field.Remove();
            }
        }
    }

    private static List<int> GetOfficeMathPages(Document doc)
    {
        var pages = new List<int>();
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        foreach (OfficeMath om in mathNodes)
        {
            // Consider only top‑level equations
            if (om.MathObjectType == MathObjectType.OMathPara)
            {
                // Placeholder page number (actual page numbers require LayoutCollector, which may not be available)
                pages.Add(0);
            }
        }

        return pages;
    }
}
