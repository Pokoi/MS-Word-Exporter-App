/*
 * File: MSWordMetaData.java
 * File Created: 29th October 2019
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

import org.w3c.dom.Document;

import word.w2004.Document2004;

public class MSWordMetaData
{
    private String title;
    private String subject;
    private String keywords;
    private String description;
    private String category;
    private String author ;
    private String lastAuthor;
    private String manager ;
    private String company;

    public MSWordMetaData(
            String title,
            String subject,
            String keywords,
            String description,
            String category,
            String author,
            String lastAuthor,
            String manager,
            String company
    )
    {
        this.title          = title;
        this.subject        = subject;
        this.keywords       = keywords;
        this.description    = description;
        this.category       = category;
        this.author         = author;
        this.lastAuthor     = lastAuthor;
        this.manager        = manager;
        this.company        = company;
    }

    public String getTitle()        { return title; }

    public String getSubject()      { return subject; }

    public String getKeywords()     { return keywords; }

    public String getDescription()  { return description; }

    public String getCategory()     { return category; }

    public String getAuthor()       { return author; }

    public String getLastAuthor()   { return lastAuthor; }

    public String getManager()      { return manager; }

    public String getCompany()      { return company; }
}
