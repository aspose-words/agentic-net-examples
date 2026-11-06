using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph before the equation.
        builder.Writeln("Paragraph before equation.");

        // Insert an EQ field that will be converted to a real OfficeMath node.
        Field field = builder.InsertField(FieldType.FieldEquation, true);

        // Write a simple EQ argument (fraction 1/2) at the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(@"\f(1,2)");

        // Update the field so that Aspose.Words converts it to OfficeMath.
        field.Update();

        // Cast the field to FieldEQ and obtain the generated OfficeMath node.
        FieldEQ fieldEq = field as FieldEQ;
        OfficeMath originalMath = fieldEq?.AsOfficeMath();

        if (originalMath == null)
            throw new InvalidOperationException("Failed to create OfficeMath from EQ field.");

        // Insert the OfficeMath node before the field start and remove the original field.
        field.Start.ParentNode.InsertBefore(originalMath, field.Start);
        field.Remove();

        // Clone the existing OfficeMath node (deep clone).
        OfficeMath clonedMath = (OfficeMath)originalMath.Clone(true);

        // Insert a new paragraph after the original equation.
        builder.Writeln("Paragraph after original equation.");

        // Locate the newly added paragraph (second paragraph in the document).
        Paragraph targetParagraph = (Paragraph)doc.GetChildNodes(NodeType.Paragraph, true)[1];

        // Append the cloned OfficeMath to the target paragraph.
        targetParagraph.AppendChild(clonedMath);

        // Save the document.
        string outputPath = "ClonedOfficeMath.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);

        // Validate that there are exactly two top‑level OfficeMath nodes.
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        int topLevelCount = 0;
        foreach (OfficeMath om in mathNodes)
        {
            if (om.MathObjectType == MathObjectType.OMathPara)
                topLevelCount++;
        }

        if (topLevelCount != 2)
            throw new InvalidOperationException($"Expected 2 top‑level OfficeMath nodes, but found {topLevelCount}.");

        // Example completed successfully.
        Console.WriteLine("Document created and OfficeMath cloned successfully.");
    }
}
