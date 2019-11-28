/*
 * File: MSWordParagraphStyle.java
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

import org.apache.poi.xwpf.usermodel.Borders;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TextAlignment;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

public class MSWordParagraphStyle
{
    ParagraphAlignment  alignment;
    Borders             borderBetween;
    Borders             borderBottom;
    Borders             borderLeft;
    Borders             borderRight;
    Borders             borderTop;
    int                 firstLineIndent;
    int                 indentationFirstLine;
    int                 indentationLeft;
    int                 indentationRight;
    int                 indentFromLeft;
    int                 indentFromRight;
    int                 spacingAfter;
    int                 spacingAfterLines;
    int                 spacingBefore;
    int                 spacingBeforeLines;
    double              spacingBetween;
    TextAlignment       verticalAlignment;



    public MSWordParagraphStyle()
    {
        this.alignment              = ParagraphAlignment.LEFT;
        this.borderBetween          = Borders.NONE;
        this.borderBottom           = Borders.NONE;
        this.borderLeft             = Borders.NONE;
        this.borderRight            = Borders.NONE;
        this.borderTop              = Borders.NONE;
        this.firstLineIndent        = 1;
        this.indentationFirstLine   = 1;
        this.indentationLeft        = 1;
        this.indentationRight       = 1;
        this.indentFromLeft         = 1;
        this.indentFromRight        = 1;
        this.spacingAfter           = 1;
        this.spacingAfterLines      = 1;
        this.spacingBefore          = 1;
        this.spacingBeforeLines     = 1;
        this.spacingBetween         = 1.0;
        this.verticalAlignment      = TextAlignment.CENTER;
    }

    public XWPFParagraph ApplyStyleOf (XWPFParagraph unstyledParagraph, XWPFParagraph styledParagraph)
    {
        unstyledParagraph.setAlignment              (styledParagraph.getAlignment());
        unstyledParagraph.setBorderBetween          (styledParagraph.getBorderBetween());
        unstyledParagraph.setBorderBottom           (styledParagraph.getBorderBottom());
        unstyledParagraph.setBorderLeft             (styledParagraph.getBorderLeft());
        unstyledParagraph.setBorderRight            (styledParagraph.getBorderRight());
        unstyledParagraph.setBorderTop              (styledParagraph.getBorderTop());
        unstyledParagraph.setFirstLineIndent        (styledParagraph.getFirstLineIndent());
        unstyledParagraph.setIndentationFirstLine   (styledParagraph.getIndentationFirstLine());
        unstyledParagraph.setIndentationLeft        (styledParagraph.getIndentationLeft());
        unstyledParagraph.setIndentationRight       (styledParagraph.getIndentationRight());
        unstyledParagraph.setIndentFromLeft         (styledParagraph.getIndentFromLeft());
        unstyledParagraph.setIndentFromRight        (styledParagraph.getIndentFromRight());
        unstyledParagraph.setSpacingAfter           (styledParagraph.getSpacingAfter());
        unstyledParagraph.setSpacingAfterLines      (styledParagraph.getSpacingAfterLines());
        unstyledParagraph.setSpacingBefore          (styledParagraph.getSpacingBefore());
        unstyledParagraph.setSpacingBeforeLines     (styledParagraph.getSpacingBeforeLines());
        unstyledParagraph.setSpacingBetween         (styledParagraph.getSpacingBetween());
        unstyledParagraph.setVerticalAlignment      (styledParagraph.getVerticalAlignment());

        return unstyledParagraph;
    }

    public XWPFParagraph ApplyDefaultStyle (XWPFParagraph unstyledParagraph)
    {
        unstyledParagraph.setAlignment              (this.alignment);
        unstyledParagraph.setBorderBetween          (this.borderBetween);
        unstyledParagraph.setBorderBottom           (this.borderBottom);
        unstyledParagraph.setBorderLeft             (this.borderLeft);
        unstyledParagraph.setBorderRight            (this.borderRight);
        unstyledParagraph.setBorderTop              (this.borderTop);
        unstyledParagraph.setFirstLineIndent        (this.firstLineIndent);
        unstyledParagraph.setIndentationFirstLine   (this.indentationFirstLine);
        unstyledParagraph.setIndentationLeft        (this.indentationLeft);
        unstyledParagraph.setIndentationRight       (this.indentationRight);
        unstyledParagraph.setIndentFromLeft         (this.indentFromLeft);
        unstyledParagraph.setIndentFromRight        (this.indentFromRight);
        unstyledParagraph.setSpacingAfter           (this.spacingAfter);
        unstyledParagraph.setSpacingAfterLines      (this.spacingAfterLines);
        unstyledParagraph.setSpacingBefore          (this.spacingBefore);
        unstyledParagraph.setSpacingBeforeLines     (this.spacingBeforeLines);
        unstyledParagraph.setSpacingBetween         (this.spacingBetween);
        unstyledParagraph.setVerticalAlignment      (this.verticalAlignment);

        return unstyledParagraph;
    }

}
