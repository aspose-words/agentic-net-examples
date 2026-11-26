using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class OfficeMathReportGenerator
{
    public static void Main()
    {
        // Create a new document and a builder for it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Helper that inserts a deterministic EQ field, converts it to OfficeMath,
        // and places the resulting OfficeMath node into the document.
        void InsertEquation(string eqSwitch)
        {
            // Insert an empty EQ field.
            Field field = builder.InsertField(FieldType.FieldEquation, true);
            FieldEQ fieldEq = field as FieldEQ;
            if (fieldEq == null)
                throw new InvalidOperationException("Failed to create FieldEQ.");

            // Move to the field separator and write the EQ argument.
            builder.MoveTo(fieldEq.Separator);
            builder.Write(eqSwitch);

            // Update the field so that Aspose.Words can parse the EQ switch.
            field.Update();

            // Convert the field to a real OfficeMath node.
            OfficeMath officeMath = fieldEq.AsOfficeMath();
            if (officeMath == null)
                throw new InvalidOperationException("EQ switch could not be converted to OfficeMath.");

            // Insert the OfficeMath node before the field start and remove the field.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            field.Remove();

            // Move the builder after the inserted OfficeMath node.
            builder.MoveTo(officeMath);
            builder.Writeln(); // Ensure a line break after the equation.
        }

        // Add some paragraphs and equations.
        builder.Writeln("First paragraph before equation.");
        InsertEquation(@"\f(1,2)"); // Simple fraction.
        builder.Writeln("Second paragraph before equation.");
        InsertEquation(@"\r(3,x)"); // Simple root.
        builder.Writeln("Third paragraph before equation.");
        InsertEquation(@"\s(5)"); // Simple superscript.

        // Save the document to a local file.
        string docPath = "Sample.docx";
        doc.Save(docPath);

        // Reload the document to follow the bootstrap rule.
        Document loadedDoc = new Document(docPath);

        // Collect all OfficeMath nodes.
        NodeCollection officeMathNodes = loadedDoc.GetChildNodes(NodeType.OfficeMath, true);

        // Prepare report lines.
        List<string> reportLines = new List<string>();
        reportLines.Add("Index,MathObjectType,ParagraphIndex");

        // Get all paragraphs for index calculation.
        NodeCollection paragraphs = loadedDoc.GetChildNodes(NodeType.Paragraph, true);

        for (int i = 0; i < officeMathNodes.Count; i++)
        {
            OfficeMath om = officeMathNodes[i] as OfficeMath;
            if (om == null)
                continue;

            // Determine the paragraph that contains this OfficeMath.
            Paragraph parentParagraph = om.ParentNode as Paragraph;
            int paragraphIndex = -1;
            if (parentParagraph != null)
                paragraphIndex = paragraphs.IndexOf(parentParagraph) + 1; // 1‑based index.

            string line = $"{i + 1},{om.MathObjectType},{paragraphIndex}";
            reportLines.Add(line);
        }

        // Write the report to a text file.
        string reportPath = "OfficeMathReport.txt";
        File.WriteAllLines(reportPath, reportLines);

        // Validate that the report file was created.
        if (!File.Exists(reportPath))
            throw new FileNotFoundException("Report file was not created.", reportPath);
    }
}
