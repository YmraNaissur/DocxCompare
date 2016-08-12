package com.naissur;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.xwpf.usermodel.Borders;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 * Paragraph creation test.
 * Creates a blank document and a paragraph with some text in it.
 * Then creates the second paragraph and applies border for it.
 */

public class ParagraphCreator {
	public static void main(String[] args) {
		
		// Creating a new document
		XWPFDocument document = new XWPFDocument();
		
		// Trying to save file
		try (FileOutputStream out = new FileOutputStream(new File("c:/tmp/createdocument.docx"))) {
			
			// Creating a new paragraph
			XWPFParagraph first_paragraph = document.createParagraph();
			
			/*
			 * Creating a run - a region of text with a common set
			 * of properties (see Apache POI javadoc)
			 */
			XWPFRun run = first_paragraph.createRun();
			run.setText("Method setText sets the text of this text run.");
			
			// Second paragraph
			XWPFParagraph second_paragraph = document.createParagraph();
			
			// Set bottom border for paragraph
			second_paragraph.setBorderBottom(Borders.BASIC_BLACK_DASHES);
			
			// Set left border for paragraph
			second_paragraph.setBorderLeft(Borders.BASIC_BLACK_DASHES);
			
			// Set right border for paragraph
			second_paragraph.setBorderRight(Borders.BASIC_BLACK_DASHES);
			
			// Set top border for paragraph
			second_paragraph.setBorderTop(Borders.BASIC_BLACK_DASHES);
			
			run = second_paragraph.createRun();
			run.setText("XWPFRun object defines a region of text with a common set of properties.");
			
			document.write(out);
			System.out.println("createdocument.docx succesfully saved");
		} catch (Exception e) {
			e.printStackTrace();
		}
		
	}
}