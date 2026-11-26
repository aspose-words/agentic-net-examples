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
        // 1. Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // 2. Insert an EQ field that will become the source OfficeMath.
        FieldEQ eqField = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write a simple fraction switch into the field separator.
        builder.MoveTo(eqField.Separator);
        builder.Write(@"\f(1,2)");

        // Ensure the field is up‑to‑date before conversion.
        eqField.Update();

        // 3. Convert the field to a real OfficeMath object.
        OfficeMath sourceMath = eqField.AsOfficeMath();
        if (sourceMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Replace the field with the OfficeMath node.
        eqField.Start.ParentNode.InsertBefore(sourceMath, eqField.Start);
        eqField.Remove();

        // 4. Add a paragraph that will precede the cloned equation.
        builder.Writeln("Clone will be inserted below:");

        // 5. Clone the existing OfficeMath node.
        OfficeMath clonedMath = (OfficeMath)sourceMath.Clone(true);
        if (clonedMath == null)
            throw new InvalidOperationException("Cloning OfficeMath failed.");

        // 6. Insert the cloned OfficeMath into a new paragraph.
        Paragraph cloneParagraph = new Paragraph(doc);
        cloneParagraph.AppendChild(clonedMath);

        // Insert the new paragraph after the placeholder paragraph.
        Paragraph placeholderParagraph = (Paragraph)builder.CurrentParagraph;
        placeholderParagraph.ParentNode.InsertAfter(cloneParagraph, placeholderParagraph);

        // 7. Save the document.
        string outputPath = "ClonedOfficeMath.docx";
        doc.Save(outputPath, SaveFormat.Docx);

        // 8. Reload and verify that two top‑level OfficeMath paragraphs exist.
        Document loadedDoc = new Document(outputPath);
        int topLevelMathCount = 0;
        foreach (Node node in loadedDoc.GetChildNodes(NodeType.OfficeMath, true))
        {
            OfficeMath om = (OfficeMath)node;
            if (om.MathObjectType == MathObjectType.OMathPara)
                topLevelMathCount++;
        }

        if (topLevelMathCount != 2)
            throw new InvalidOperationException($"Expected 2 top‑level OfficeMath nodes, but found {topLevelMathCount}.");
    }
}
