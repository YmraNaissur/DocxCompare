package com.naissur;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

/**
 * Docx creation test
 * Creating blank document
 */

public class BlankDocumentCreator {
	public static void main(String[] args) {
		
		// Creating Blank Document
		XWPFDocument document = new XWPFDocument();
		
		// Write the Document in file system
		try (FileOutputStream out = new FileOutputStream(new File("c:/tmp/createdocument.docx"))) {
			
			document.write(out);
			System.out.println("Createdocument.docx written successfully.");
			
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}