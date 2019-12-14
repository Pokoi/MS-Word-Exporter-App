/*
 * File: MSWordTemplate.java
 * File Created: 14th December 2019
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

import android.content.Context;
import android.content.Intent;
import android.net.Uri;
import android.os.StrictMode;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.PrintWriter;

import word.w2004.Document2004;

public class MSWordTemplate
{
    private String content;

    public String GetContent() { return this.content;}

    public MSWordTemplate(String path)
    {
        try {
            content = GetStringFromFile(path);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    ////////////////////////////////////////////////////////////////////////////////////////////////
    // READING THE FILE
    ////////////////////////////////////////////////////////////////////////////////////////////////

    //http://www.java2s.com/Code/Java/File-Input-Output/ConvertInputStreamtoString.htm
    public String GetStringFromFile (String filePath) throws Exception {
        File fl = new File(filePath);
        FileInputStream fin = new FileInputStream(fl);
        String ret = ConvertStreamToString(fin);
        fin.close();
        return ret;
    }


    private String ConvertStreamToString(InputStream is) throws Exception {
        BufferedReader reader = new BufferedReader(new InputStreamReader(is));
        StringBuilder sb = new StringBuilder();
        String line = null;
        while ((line = reader.readLine()) != null) {
            sb.append(line).append("\n");
        }
        reader.close();
        return sb.toString();
    }

    ////////////////////////////////////////////////////////////////////////////////////////////////
    // REPLACING CONTENT
    ////////////////////////////////////////////////////////////////////////////////////////////////

    public void Replace(String placeHolder, String newValue)
    {
        this.content = this.content.replace(placeHolder, newValue);
    }

    public void Extract (String start, String end)
    {
        this.content = this.content.substring(this.content.indexOf(start), this.content.indexOf(end) + end.length());
        Replace(start, "");
        Replace(end, "");
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
}
