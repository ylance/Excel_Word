package com.cm.oe.budget.gen;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.Borders;
import org.apache.poi.xwpf.usermodel.BreakClear;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TextAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableCell.XWPFVertAlign;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.impl.xb.xmlschema.SpaceAttribute;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFldChar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHdrFtr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STFldCharType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHdrFtr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.impl.CTHdrFtrImpl;

public class ymtest {
	public static void main(String[] args) throws Exception {
		ymtest t=new ymtest();
		String url = "testfiles/testheader.docx";
		InputStream is = new FileInputStream("testfiles/testheader.docx");
		XWPFDocument document = new XWPFDocument(is);
		t.simpleDateHeader2(document);
		t.saveDocument(document, "D:/sys_"+ System.currentTimeMillis() + ".docx");
	}
	
	
	public void simpleDateHeader2(XWPFDocument document) throws Exception {
		CTHdrFtr headerFooter = CTHdrFtr.Factory.newInstance();
        CTP ctp = headerFooter.addNewP();
		XWPFParagraph codePara = new XWPFParagraph(ctp, document);
		XWPFRun r1 = codePara.createRun();

		r1 = codePara.createRun();
		CTRPr rpr = r1.getCTR().isSetRPr() ? r1.getCTR().getRPr() : r1.getCTR().addNewRPr();
		CTFonts fonts = rpr.isSetRFonts() ? rpr.getRFonts() : rpr.addNewRFonts();
		r1 = codePara.createRun();
		r1.setText("中国移动通信");
		r1.setFontSize(11);
		rpr = r1.getCTR().isSetRPr() ? r1.getCTR().getRPr() : r1.getCTR().addNewRPr();
		fonts = rpr.isSetRFonts() ? rpr.getRFonts() : rpr.addNewRFonts();
		fonts.setAscii("微软雅黑");
		fonts.setEastAsia("微软雅黑");
		fonts.setHAnsi("微软雅黑");
		
		codePara.setAlignment(ParagraphAlignment.CENTER);
		codePara.setVerticalAlignment(TextAlignment.CENTER);
		codePara.setBorderBottom(Borders.THICK);
		XWPFParagraph[] newparagraphs = new XWPFParagraph[1];
		newparagraphs[0] = codePara;
		List<XWPFHeader> headers=document.getHeaderList();
		for (  XWPFHeader header : headers) {
			if(header.getPackagePart().getPartName().toString().equals("/word/header1.xml")){
				continue;
			}			
			if(header.getPackagePart().getPartName().toString().equals("/word/header2.xml")){
				continue;
			}
			System.out.println(header.getPackagePart().getPartName());
			header.setHeaderFooter(headerFooter);
		}
	}

	public void saveDocument(XWPFDocument document,String savePath) throws Exception{
		FileOutputStream fos = new FileOutputStream(savePath);
		document.write(fos);
		fos.close();
	}
	
	
}