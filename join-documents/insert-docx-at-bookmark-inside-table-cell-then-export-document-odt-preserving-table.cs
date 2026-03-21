using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsInsertDocAtBookmark
{
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create the destination document and a table that contains a bookmark.
            // -----------------------------------------------------------------
            Document dstDoc = new Document();                     // create blank document
            DocumentBuilder builder = new DocumentBuilder(dstDoc); // builder for the destination

            // Build a simple 1x1 table.
            builder.StartTable();          // <TABLE>
            builder.InsertCell();          //   <ROW><CELL>

            // Insert a bookmark inside the cell where the source document will be placed.
            builder.StartBookmark("InsertHere");
            builder.EndBookmark("InsertHere");

            builder.EndRow();              //   </ROW>
            builder.EndTable();            // </TABLE>

            // -----------------------------------------------------------------
            // 2. Create the source document that we want to insert.
            // -----------------------------------------------------------------
            Document srcDoc = new Document();                     // create a new source document
            DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
            srcBuilder.Writeln("Hello from the source document!"); // add some content

            // -----------------------------------------------------------------
            // 3. Move the builder's cursor to the bookmark and insert the source document.
            // -----------------------------------------------------------------
            builder.MoveToBookmark("InsertHere"); // position cursor at the bookmark
            // InsertDocument copies the whole source document at the current cursor position.
            // KeepSourceFormatting preserves the original formatting of the source.
            builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            // -----------------------------------------------------------------
            // 4. Save the resulting document as ODT, preserving the table structure.
            // -----------------------------------------------------------------
            OdtSaveOptions odtOptions = new OdtSaveOptions(); // default ODT options
            dstDoc.Save("Result.odt", odtOptions);            // export to ODT
        }
    }
}
