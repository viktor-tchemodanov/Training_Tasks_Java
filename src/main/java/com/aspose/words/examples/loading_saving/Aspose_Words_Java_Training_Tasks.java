package com.aspose.words.examples.loading_saving;
import com.aspose.words.*;
import com.aspose.words.Document;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.examples.Utils;

public class Aspose_Words_Java_Training_Tasks {

    // Extract the path to the documents directory.

    static String dataDir = Utils.getDataDir(Aspose_Words_Java_Training_Tasks.class);

    public static void main(String[] args) throws Exception {

        //ExStart:
        Task1_HelloWorldConsole();
        Task2_DocToPdf();
        Task3_HelloWorldAW();
        Task4_JoinTwoDocuments();
        Task5_ReplaceTextOfBookmark();

        System.out.println("\nTasks completed.\nFiles saved at " + dataDir);
        System.out.println("\n\nProgram Finished. Press any key to exit....");
        System.in.read();
    }

    // Tasks 1
    // Printing "Hello World!".
    public static void Task1_HelloWorldConsole() {

       System.out.println("Hello World!");
    }

    // Tasks 2
    // Converting a document from Doc to Pdf format.
    public static void Task2_DocToPdf() throws Exception {

        Document doc = new Document(dataDir + "Tech_Article_Viktor.doc");

        // Save the document in PDF format.
             doc.save(dataDir + "Tech_Article_Viktor_out.pdf");
    }

    // Task 3
    // Programmatically create a document and insert a paragraph with text Hello World and save to disk.
    public static void Task3_HelloWorldAW() throws Exception {

    // Create a blank document.
        String fileName = "CreateDocument_out.docx";
        Document doc = new Document();

     // Use a document builder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write("Hello World!");

     // Save the finished document to disk.
        doc.save(dataDir + fileName);

    }

    // Task 4
    // Programmatically join two documents and save to disk.
    public static void Task4_JoinTwoDocuments() throws Exception {

        Document dstDoc = new Document(dataDir + "TestFile.Destination.doc");

        Document srcDoc = new Document(dataDir + "TestFile.Source.doc");

     // Append the source document using the original styles found in the source document.
        dstDoc.appendDocument(srcDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);

     // Save the new document to disk.
        dstDoc.save(dataDir + "Joined_Document_out.doc");

    }

    // Tasks 5
    // Programmatically open the document with a bookmark, replace the bookmarked text, and save the new document to disk.
        public static void Task5_ReplaceTextOfBookmark() throws Exception {

        Document doc = new Document(dataDir + "Document_w_bookmark.doc");

     // Use the indexer of the Bookmarks collection to obtain the desired bookmark.
        Bookmark bookmark = doc.getRange().getBookmarks().get("Characteristics");
     // Get the name and text of the bookmark.
        String name = bookmark.getName();
        String text = bookmark.getText();

     // Set the bookmark text.
        bookmark.setText("This is a new bookmarked text.");

     // Save the new document to disk.
        doc.save(dataDir + "Document_new_bookmark_out.doc");

    }

}