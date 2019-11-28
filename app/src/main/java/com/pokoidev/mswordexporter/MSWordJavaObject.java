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

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.File;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;


public class MSWordJavaObject
{

    XWPFDocument         document;
    MSWordParagraphStyle paragraphStyler;

    public MSWordJavaObject()
    {
        this.document        = new XWPFDocument();
        this.paragraphStyler = new MSWordParagraphStyle();
    }

    public MSWordJavaObject(String fileName) throws Exception
    {
        Path documentPath   = Paths.get(fileName);
        this.document       = new XWPFDocument(Files.newInputStream(documentPath));
        document.close();

        this.paragraphStyler = new MSWordParagraphStyle();
    }


    public void AddParagraph(String paragraphText)
    {
        XWPFParagraph newParagraph  = document.createParagraph();
        XWPFRun paragraphRun        = newParagraph.createRun();

        paragraphRun.setText(paragraphText);
    }

    public void AddImage (String imageName, int width, int height) throws Exception {
        XWPFParagraph image = document.createParagraph();
        XWPFRun imageRun    = image.createRun();

        Path imagePath = Paths.get(ClassLoader.getSystemResource(imageName).toURI());
        imageRun.addPicture(Files.newInputStream(imagePath), XWPFDocument.PICTURE_TYPE_PNG, imagePath.getFileName().toString(), width, height);
    }

    public void ExportFile(File file) throws Exception
    {
        FileOutputStream out = new FileOutputStream(file);
        document.write(out);
        out.close();
        document.close();
    }



}

