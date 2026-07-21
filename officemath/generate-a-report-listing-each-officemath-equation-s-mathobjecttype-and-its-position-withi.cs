using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class OfficeMathReportExample
{
    public static void Main()
    {
        // Paths for output files
        string docPath = "OfficeMathDocument.docx";
        string reportPath = "OfficeMathReport.txt";

        // Create a new blank document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Simple EQ strings to create deterministic OfficeMath equations
        string[] eqStrings = new string[]
        {
            @"\f(1,2)",          // Fraction 1/2
            @"\r(3,x)",          // Cube root of x
            @"\i \su(n=1,5,n)", // Integral with summation
            @"\s \up8(Sup)",    // Superscript
            @"\s \do8(Sub)"     // Subscript
        };

        // Insert each equation into its own paragraph using the EQ-field bootstrap workflow
        foreach (string eq in eqStrings)
        {
            // Insert an EQ field
            FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

            // Write the EQ arguments into the field separator
            builder.MoveTo(field.Separator);
            builder.Write(eq);

            // Return the builder to the field start's parent (the paragraph) to continue building
            builder.MoveTo(field.Start.ParentNode);

            // Convert the field to a real OfficeMath object
            OfficeMath officeMath = field.AsOfficeMath();
            if (officeMath != null)
            {
                // Insert the OfficeMath node before the field start
                field.Start.ParentNode.InsertBefore(officeMath, field.Start);
                // Remove the original field
                field.Remove();
            }

            // Start a new paragraph for the next equation
            builder.InsertParagraph();
        }

        // Save the document containing the equations
        doc.Save(docPath);

        // Generate a report listing each OfficeMath equation's MathObjectType and its position
        NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

        using (StreamWriter writer = new StreamWriter(reportPath))
        {
            writer.WriteLine($"Total OfficeMath equations: {officeMathNodes.Count}");
            writer.WriteLine();

            for (int i = 0; i < officeMathNodes.Count; i++)
            {
                OfficeMath om = (OfficeMath)officeMathNodes[i];
                MathObjectType mathType = om.MathObjectType;

                // Determine the paragraph and section that contain this OfficeMath node
                Paragraph paragraph = om.ParentParagraph;
                Section section = (Section)paragraph.GetAncestor(NodeType.Section);

                int sectionIndex = doc.Sections.IndexOf(section);
                int paragraphIndex = section.Body.Paragraphs.IndexOf(paragraph);

                writer.WriteLine($"Equation {i + 1}:");
                writer.WriteLine($"  MathObjectType : {mathType}");
                writer.WriteLine($"  Section Index  : {sectionIndex}");
                writer.WriteLine($"  Paragraph Index: {paragraphIndex}");
                writer.WriteLine();
            }
        }

        // Validate that the report file was created
        if (!File.Exists(reportPath))
            throw new InvalidOperationException("Report file was not created.");

        // Optional: write a short confirmation to the console
        Console.WriteLine($"Document saved to '{docPath}'.");
        Console.WriteLine($"Report generated at '{reportPath}'.");
    }
}
