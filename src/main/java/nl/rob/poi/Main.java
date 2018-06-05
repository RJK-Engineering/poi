package nl.rob.poi;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class Main {

	static String file = "c:\\temp\\a.docx";
	static String outfile = "c:\\temp\\a_out.docx";
	static HashMap<String, String> tags;
	
	public static void main(String[] args) throws FileNotFoundException, IOException {
		
		tags = new HashMap<String, String>();
		tags.put("TAG", "tag");
		tags.put("TAG2", "tag2");
		tags.put("BOLD ITALIC", "bold italic");
		

		XWPFDocument doc = new XWPFDocument(
			new FileInputStream(file)
		);

		processBody(doc);
		
		FileOutputStream out = new FileOutputStream(outfile);
		doc.write(out);
		doc.close();
	}

	public static void processBody(XWPFDocument doc) {
		for (XWPFTable table : doc.getTables())
			for (XWPFTableRow row : table.getRows())
				for (XWPFTableCell cell : row.getTableCells()) 
					for (XWPFParagraph p : cell.getParagraphs())
						processParagraph(p);
		for (XWPFParagraph p : doc.getParagraphs())
			processParagraph(p);
	}
		
	private static void processParagraph(XWPFParagraph p) {
		String txt = "";
		XWPFRun start = null;
		for (XWPFRun r : p.getRuns()) {
			String text = r.text(); 
			if (text != null) {
				if (text.contains("<<"))
					start = r;
				if (start != null) {
					txt += text; 
					r.setText("");
				}
				if (text.contains(">>")) {
					txt = replaceTags(start, txt);
					start.setText(txt);
					txt = "";
					start = null;
				}
			}
		}
	}

	private static String replaceTags(XWPFRun r, String text) {
		Pattern pattern = Pattern.compile("<<(.+?)>>");
		Matcher matcher = pattern.matcher(text);
		while (matcher.find()) {
			String tag = matcher.group(1);
			text = text.replaceFirst("<<" + tag + ">>", getReplacement(tag));
		    System.out.println("tag: " + tag + " text: " + text);
		}			
		return text;
	}

	private static String getReplacement(String tag) {
		String text = tags.get(tag);
		if (text == null)
			return "<<" + tag + ">>";
		else
			return text;
	}

}
