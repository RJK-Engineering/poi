package nl.rob.poi;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import nl.novadoc.utils.Logger;

public class TagReplacer {

	static Logger log = Logger.getLogger();

	private HashMap<String, String> tags;
	
	public TagReplacer() {
		tags = new HashMap<String, String>();
	}

	public TagReplacer(HashMap<String, String> tags) {
		this.tags = tags;
	}

	public void addTag(String tag, String replacement) {
		tags.put(tag, replacement);
	}

	public void processBody(XWPFDocument doc) {
		for (XWPFTable table : doc.getTables())
			for (XWPFTableRow row : table.getRows())
				for (XWPFTableCell cell : row.getTableCells()) 
					for (XWPFParagraph p : cell.getParagraphs())
						processParagraph(p);
		for (XWPFParagraph p : doc.getParagraphs())
			processParagraph(p);
	}
		
	private void processParagraph(XWPFParagraph p) {
		String txt = "";
		ArrayList<XWPFRun> runs = new ArrayList<XWPFRun>();
		for (XWPFRun r : p.getRuns()) {
			String text = r.getText(0); 
			if (text == null)
				continue;
			if (runs.size() == 0 && ! text.contains("<<"))
				continue;

			txt += text;
			runs.add(r);
			r.setText("");

			if (text.contains(">>")) {
				updateRuns(runs, txt);
				txt = "";
				runs.clear();
			}
		}
	}

	private void updateRuns(ArrayList<XWPFRun> runs, String txt) {
		txt = replaceTags(runs.get(0), txt);
		
		int runTxtLength = txt.length() / runs.size() + 1;
		String[] runTxts = txt.split("(?<=\\G.{" + runTxtLength + "})");
		
		log.debug(txt);
		log.debug(txt.length() + " " + runs.size() + " " + runTxtLength);
		log.debug(Arrays.toString(runTxts));
		
		for (int i=0; i<runs.size(); i++) {
			runs.get(i).setText(runTxts[i], 0);
		}		
	}

	private String replaceTags(XWPFRun r, String text) {
		Pattern pattern = Pattern.compile("<<(.+?)>>");
		Matcher matcher = pattern.matcher(text);
		while (matcher.find()) {
			String tag = matcher.group(1);
			text = text.replaceFirst("<<" + tag + ">>", getReplacement(tag));
		    System.out.println("tag: " + tag + " text: " + text);
		}			
		return text;
	}

	private String getReplacement(String tag) {
		String text = tags.get(tag);
		if (text == null)
			return "<<" + tag + ">>";
		else
			return text;
	}

}
