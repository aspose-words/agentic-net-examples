using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Original MathML string (metadata only)
        string mathML = "<math xmlns='http://www.w3.org/1998/Math/MathML'><mfrac><mi>a</mi><mi>b</mi></mfrac></math>";

        // Create a new document and a builder
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph before the equation
        builder.Writeln("This paragraph will contain an equation inserted from MathML.");

        // Insert a comment that stores the original MathML for reference
        Comment comment = new Comment(doc, "Author", "AU", DateTime.Now);
        comment.Paragraphs.Add(new Paragraph(doc));
        comment.FirstParagraph.AppendChild(new Run(doc, $"Original MathML: {mathML}"));
        builder.CurrentParagraph.AppendChild(comment);

        // Insert an EQ field (deterministic bootstrap for OfficeMath)
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        FieldEQ fieldEq = field as FieldEQ;
        if (fieldEq == null)
            throw new InvalidOperationException("Failed to create FieldEQ.");

        // Write a simple EQ argument (fraction 1 over 2) into the field separator
        builder.MoveTo(fieldEq.Separator);
        builder.Write(@"\f(1,2)");

        // Update the field so that Aspose.Words can convert it to OfficeMath
        fieldEq.Update();

        // Convert the field to a real OfficeMath node
        OfficeMath officeMath = fieldEq.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("EQ field could not be converted to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the original field
        Node fieldStart = fieldEq.Start;
        fieldStart.ParentNode.InsertBefore(officeMath, fieldStart);
        fieldEq.Remove();

        // Add a paragraph after the equation
        builder.Writeln();
        builder.Writeln("Paragraph after the equation.");

        // Save the document
        string outputPath = "Output.docx";
        doc.Save(outputPath, SaveFormat.Docx);

        // Verify the output file exists
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);

        // Verify that at least one top‑level OfficeMath node exists
        NodeCollection mathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);
        if (mathNodes.Count == 0)
            throw new InvalidOperationException("No OfficeMath nodes were found in the saved document.");

        // Write a simple report
        string reportPath = "Report.txt";
        File.WriteAllText(reportPath, $"Inserted OfficeMath count: {mathNodes.Count}");
        if (!File.Exists(reportPath))
            throw new FileNotFoundException("Report file was not created.", reportPath);
    }
}
