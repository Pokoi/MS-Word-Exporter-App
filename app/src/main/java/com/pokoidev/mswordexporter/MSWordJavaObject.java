/*
 * File: MSWordJavaObject.java
 * File Created: 27th October 2019
 * ––––––––––––––––––––––––
 * Author: Jesus Fermin, 'Pokoi', Villar  (hello@pokoidev.com)
 * ––––––––––––––––––––––––
 * MIT License
 *
 * Copyright (c) 2019 Pokoidev
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of
 * this software and associated documentation files (the "Software"), to deal in
 * the Software without restriction, including without limitation the rights to
 * use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies
 * of the Software, and to permit persons to whom the Software is furnished to do
 * so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */

package com.pokoidev.mswordexporter;

import java.io.File;
import java.io.PrintWriter;

import word.api.interfaces.IDocument;
import word.w2004.Document2004;
import word.w2004.elements.BreakLine;
import word.w2004.elements.Heading1;
import word.w2004.elements.Heading2;
import word.w2004.elements.Heading3;
import word.w2004.elements.Image;
import word.w2004.elements.Paragraph;
import word.w2004.elements.Table;
import word.w2004.elements.tableElements.TableEle;


public class MSWordJavaObject
{
    IDocument document;
    MSWordMetaData  metaData;

    /**
     * Creates a MSWordJavaObject object
     * @param metaData The metadata info of the new document
     */
    public MSWordJavaObject(MSWordMetaData metaData)
    {
        this.metaData = metaData;

        CreateDocument();
        ApplyMetadata();
    }

    /**
     * Gets the metadata info of the Microsoft Word document
     * @return Metadata info of the Microsoft Word document
     */
    public MSWordMetaData getMetaData() { return this.metaData;}

    ////////////////////////////////////////////////////////////////////////////////////////////////
    // ADDING ELEMENTS TO THE DOCUMENT BODY CONTENT
    ////////////////////////////////////////////////////////////////////////////////////////////////

    /**
     * Adds a new paragraph to the document with the given text
     * @param paragraphText The new paragraph's text
     */
    public void AddParagraph(String paragraphText)
    {
        document.addEle(Paragraph.with(paragraphText).create());
    }

    /**
     * Adds a new header text with the given text
     * @param headingText The new heading's text
     * @param headingLevel Level of heading (1, 2 or 3)
     */
    public void AddHeading(String headingText, int headingLevel)
    {
        switch (headingLevel)
        {
            case 1:
                document.addEle(Heading1.with(headingText).create());
                break;
            case 2:
                document.addEle(Heading2.with(headingText).create());
                break;
            case 3:
                document.addEle(Heading3.with(headingText).create());
                break;
        }
    }

    /**
     * Adds an image from an URL to the Microsoft Word document
     * @param imageURL The image URL
     * @param width The desired width of the image
     * @param height The desired height of the image
     */
    public void AddImageFromWeb (String imageURL, int width, int height)
    {
        this.document.addEle(   Image.from_WEB_URL(imageURL)
                .setHeight(Integer.toString(height))
                .setWidth(Integer.toString(width))
                .create()
        );
    }

    /**
     * Adds an image from a path to the Microsoft Word document
     * @param imagePath The image path
     * @param width The desired width of the image
     * @param height The desired height of the image
     */
    public void AddImageFromFullPath (String imagePath, int width, int height)
    {
        this.document.addEle(   Image.from_FULL_LOCAL_PATHL(imagePath)
                .setHeight(Integer.toString(height))
                .setWidth(Integer.toString(width))
                .create()
        );
    }

    /**
     * Adds a new blank line for the given times
     * @param times The amount of new blank lines
     */
    public void AddBlankLines(int times)
    {
        this.document.addEle(BreakLine.times(times).create());
    }

    /**
     * Creates a table for the Microsoft Word document
     * @return The created table
     */
    public Table CreateTable()
    {
        Table table = new Table();
        table.setRepeatTableHeaderOnEveryPage();
        return table;
    }

    /**
     * Adds a new row in the given table with the given info
     * @param table The table where insert the new row
     * @param columnsInfo The strings with the info of each column in this row
     * @return The given table with the inserted row
     */
    public Table AddRowToTable(Table table, String... columnsInfo)
    {
        table.addTableEle(TableEle.TD, columnsInfo);
        return table;
    }

    /**
     * Closes a given table and adds it to the Microsoft Word document
     * @param table The table to insert
     */
    public void CloseTable(Table table)
    {
        this.document.addEle(table);
    }

    ////////////////////////////////////////////////////////////////////////////////////////////////
    // ADDING ELEMENTS TO THE HEADER
    ////////////////////////////////////////////////////////////////////////////////////////////////

    /**
     * Adds the given text to the header of the Microsoft Word document
     * @param headerText
     */
    public void AddTextToHeader(String headerText)
    {
        this.document.getHeader().addEle(Paragraph.with(headerText).create());
    }

    /**
     * Adds an image from an URL to the header of the Microsoft Word document
     * @param imageURL The image URL
     * @param width The desired width of the image
     * @param height The desired height of the image
     */
    public void AddImageToHeaderFromWeb(String imageURL, int width, int height)
    {
        this.document.getHeader().addEle(   Image.from_WEB_URL(imageURL)
                .setHeight(Integer.toString(height))
                .setWidth(Integer.toString(width))
                .create()
        );
    }

    /**
     * Adds an image from path to the header of the Microsoft Word document
     * @param imagePath The image path
     * @param width The desired width of the image
     * @param height The desired height of the image
     */
    public void AddImageToHeaderFromFullPath(String imagePath, int width, int height)
    {
        this.document.getHeader().addEle(   Image.from_FULL_LOCAL_PATHL(imagePath)
                .setHeight(Integer.toString(height))
                .setWidth(Integer.toString(width))
                .create()
        );
    }

    ////////////////////////////////////////////////////////////////////////////////////////////////
    // ADDING ELEMENTS TO THE FOOTER
    ////////////////////////////////////////////////////////////////////////////////////////////////

    /**
     * Adds the given text to the footer of the Microsoft Word document
     * @param footerText
     */
    public void AddTextToFooter(String footerText)
    {
        this.document.getFooter().addEle(Paragraph.with(footerText).create());
    }

    /**
     * Adds an image from an URL to the footer of the Microsoft Word document
     * @param imageURL The image URL
     * @param width The desired width of the image
     * @param height The desired height of the image
     */
    public void AddImageToFooterFromWeb (String imageURL, int width, int height)
    {
        this.document.getFooter().addEle(   Image.from_WEB_URL(imageURL)
                .setHeight(Integer.toString(height))
                .setWidth(Integer.toString(width))
                .create()
        );
    }

    /**
     * Adds an image from path to the footer of the Microsoft Word document
     * @param imagePath The image path
     * @param width The desired width of the image
     * @param height The desired height of the image
     */
    public void AddImageToFooterFromFullPath(String imagePath, int width, int height)
    {
        this.document.getFooter().addEle(   Image.from_FULL_LOCAL_PATHL(imagePath)
                .setHeight(Integer.toString(height))
                .setWidth(Integer.toString(width))
                .create()
        );
    }

    ////////////////////////////////////////////////////////////////////////////////////////////////
    // DOCUMENT OPEN/CLOSE
    ////////////////////////////////////////////////////////////////////////////////////////////////

    /**
     * Closes and exports the document, saving it on the local storage
     * @param file The file to export
     * @throws Exception
     */
    public void ExportFile(File file) throws Exception
    {
        String content      = document.getContent();
        PrintWriter writer  = new PrintWriter(file);

        writer.println(content);
        writer.close();
    }

    /**
     * Sends an email with the generated document attached
     * @param file The file to attach
     * @param emailDirections The email directions to send to
     * @param emailSubject The subject of the mail to send
     * @param context A context reference to start the send activity
     * @throws Exception
     */
    public void SendToMail(File file, String[] emailDirections, String emailSubject, Context context) throws Exception
    {
        StrictMode.VmPolicy.Builder builder = new StrictMode.VmPolicy.Builder();
        StrictMode.setVmPolicy(builder.build());

        ExportFile(file);


        Intent emailIntent = new Intent (Intent.ACTION_SEND);

        emailIntent.setType("text/plain");
        emailIntent.putExtra(Intent.EXTRA_EMAIL, emailDirections);
        emailIntent.putExtra(Intent.EXTRA_SUBJECT, emailSubject);

        Uri uri = Uri.fromFile(file);
        emailIntent.putExtra(Intent.EXTRA_STREAM, uri);
        context.startActivity(Intent.createChooser(emailIntent, "Pick an Email provider"));

    }

    /**
     * Document initialization
     */
    private void CreateDocument()
    {
        this.document = new Document2004();
        this.document.encoding(Document2004.Encoding.UTF_8);
    }

    /**
     * Apply the metadata to the Microsoft Word document
     */
    private void ApplyMetadata ()
    {
        this.document.title(this.metaData.getTitle());
        this.document.subject(this.metaData.getSubject());
        this.document.keywords(this.metaData.getKeywords());
        this.document.description(this.metaData.getDescription());
        this.document.category(this.metaData.getCategory());
        this.document.author(this.metaData.getAuthor());
        this.document.lastAuthor(this.metaData.getLastAuthor());
        this.document.manager(this.metaData.getManager());
        this.document.company(this.metaData.getCompany());
    }
}

