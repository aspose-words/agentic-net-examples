using System;
using Aspose.Words;

namespace AsposeWordsInsertDemo
{
    class Program
    {
        static void Main()
        {
            // Load the template document that contains a bookmark named "InsertHere".
            Document template = new Document("Template.docx");

            // Load the document that will be inserted into the template.
            Document docToInsert = new Document("Insert.docx");

            // Create a DocumentBuilder attached to the template.
            DocumentBuilder builder = new DocumentBuilder(template);

            // Move the cursor to the bookmark where the insertion should occur.
            builder.MoveToBookmark("InsertHere");

            // Insert the source document at the bookmark position.
            // KeepSourceFormatting preserves the original formatting of the inserted document.
            builder.InsertDocument(docToInsert, ImportFormatMode.KeepSourceFormatting);

            // Save the resulting document.
            template.Save("Result.docx");
        }
    }
}
