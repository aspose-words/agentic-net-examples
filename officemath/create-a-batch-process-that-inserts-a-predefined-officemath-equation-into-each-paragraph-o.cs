using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class OfficeMathBatchInsert
{
    public static void Main()
    {
        // Create a sample document with several paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"This is paragraph {i}.");
        }

        // Save the sample input document.
        string inputPath = "Input.docx";
        doc.Save(inputPath, SaveFormat.Docx);

        // Load the document for batch processing.
        Document processedDoc = new Document(inputPath);
        DocumentBuilder procBuilder = new DocumentBuilder(processedDoc);

        // Retrieve all paragraph nodes.
        NodeCollection paragraphs = processedDoc.GetChildNodes(NodeType.Paragraph, true);

        // Predefined EQ argument for the OfficeMath equation.
        const string eqArgument = @"\f(1,2)";

        foreach (Paragraph para in paragraphs)
        {
            // Position the builder at the start of the current paragraph.
            procBuilder.MoveTo(para);

            // Insert an EQ field which will be converted to OfficeMath.
            Field field = procBuilder.InsertField(FieldType.FieldEquation, true);
            FieldEQ fieldEq = field as FieldEQ;
            if (fieldEq == null)
                continue; // Safety check; should not occur.

            // Write the EQ argument into the field separator.
            procBuilder.MoveTo(fieldEq.Separator);
            procBuilder.Write(eqArgument);

            // Convert the field to a real OfficeMath node.
            OfficeMath officeMath = fieldEq.AsOfficeMath();
            if (officeMath != null)
            {
                // Insert the OfficeMath node before the field start node.
                CompositeNode parent = field.Start.ParentNode as CompositeNode;
                if (parent != null)
                {
                    parent.InsertBefore(officeMath, field.Start);
                }

                // Remove the original field, leaving only the OfficeMath node.
                field.Remove();
            }
        }

        // Save the processed document.
        string outputPath = "Output.docx";
        processedDoc.Save(outputPath, SaveFormat.Docx);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");
    }
}
