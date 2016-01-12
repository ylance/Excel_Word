package com.cm.oe.test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Table;

public class GetTableTest {
	ReadWord rw = new ReadWord();
	public static void main(String[] args) {
		GetTableTest gt = new GetTableTest();
		try {
			gt.getTables();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	public void getTables() throws IOException{
		String wordPath = "testfiles/template.doc";
		FileInputStream fisw = new FileInputStream(wordPath);
		HWPFDocument doc = new HWPFDocument(fisw);
		Range range = rw.getRange(doc);
		System.out.println(range.numParagraphs());
		Paragraph paragraph = range.getParagraph(2);
		System.out.println(paragraph.text());
		//Table t = range.getTable(paragraph);
		//System.out.println(t.getRow(1));
	}
}
