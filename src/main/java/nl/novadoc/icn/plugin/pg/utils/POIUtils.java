package nl.novadoc.icn.plugin.pg.utils;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import com.filenet.api.collection.ContentElementList;
import com.filenet.api.core.ContentTransfer;
import com.filenet.api.core.Document;
import com.filenet.api.core.Folder;

import nl.novadoc.icn.plugin.pg.utils.helper.ReplaceHelper;
import nl.novadoc.utils.Logger;

/**
 * @author jrwerkman
 */
public class POIUtils {
	static Logger log = Logger.getLogger();
	private static final String TAG = "POIUtils";
	

	public static XWPFDocument replaceAllPlaceholders(String placeholders, String replacePropertiesString, 
			Document doc, Folder folder) throws Exception {
		String[] values = placeholders.split(",");
		// Get values of document property, String will be empty when the propery was not found or unsupported.
		String[] replaceProperties = replacePropertiesString.split(",");
		String[] replaceValues = new String[values.length];
		
		for(int i=0; i<replaceProperties.length; i++) {
			if(replaceProperties[i].startsWith("C:"))
				replaceValues[i] = FileNetUtils.getObjectProperty(folder, replaceProperties[i].replaceFirst("C:", ""));
			else if(replaceProperties[i].startsWith("D:"))
				replaceValues[i] = FileNetUtils.getObjectProperty(doc, replaceProperties[i].replaceFirst("D:", ""));
			else
				throw new Exception("Found an invalid property: " + replaceProperties[i] + ". Thje property needs to start with C: or D:");
		}
		
		XWPFDocument poiDoc = POIUtils.convertFNDocToXWPFDoc(doc);
		POIUtils.replaceTextParagraphs(poiDoc, values, replaceValues);
		POIUtils.replaceTextTables(poiDoc, values, replaceValues);
		
		return poiDoc;
	}
	
	
	/**
	 * Creates a POI doc from a (docx) document in filenet
	 *  
	 * @param fnDoc
	 * @return
	 * @throws IOException
	 */
	public static XWPFDocument convertFNDocToXWPFDoc(Document fnDoc) throws Exception {
		log.debug(TAG + ".convertFNDocToXWPFDoc");

		XWPFDocument poiDoc = null;
		
		try {
			ContentElementList cel = fnDoc.get_ContentElements();
			InputStream stream = null;
	
			if (cel.size() > 1) {
				log.error("Document: Containt more then one element!");
			} else if (cel.size() < 1) {
				log.error("Unable to save Document: No elements found!!");
			} else {
				@SuppressWarnings("unchecked")
				final Iterator<ContentTransfer> iter = cel.iterator();
	
				while (iter.hasNext()) {
					ContentTransfer ct = iter.next();
					stream = ct.accessContentStream();
				}
			}
			poiDoc = new XWPFDocument(stream);
		} catch(Exception e) {
			log.error("Could not convert FN doc to XWPFDocument");
			log.error(e.getMessage());
			throw e;
		}
		return poiDoc;
	}
	
	/**
	 * Writes a POT docx to an path with a certain filename
	 * 
	 * @param doc
	 * @param path
	 * @param filename
	 * @throws Exception
	 */
	public static void writeXWPFDocument(XWPFDocument doc, String path, String filename) throws Exception {
		log.debug(TAG + ".writeXWPFDocument");

		try {
			File f = new File(path, filename);
			doc.write(new FileOutputStream(f));
		} catch (FileNotFoundException e) {
			log.error("Could not find path to write the file: " + path);
			throw e;
		} catch (IOException e) {
			log.error("Could not write the file " + filename + " to path " + path);
			throw e;
		}
	}
	
	
	/**
	 * Replaces a string value (or more) in paragraphs, by another String value
	 * 
	 * @param xwpfDoc
	 * @param value
	 * @param replaceValue
	 * @param onlyFirstOccurrence
	 */
	public static void replaceTextParagraphs(XWPFDocument xwpfDoc, String[] values, String[] replaceValues) throws Exception {
		log.debug(TAG + ".replaceTextParagraphs");
		
		if(values.length != replaceValues.length) {
			log.error("There are more values to replace than values to replace them with.");
			return;
		}

		try {
			for (XWPFParagraph p : xwpfDoc.getParagraphs())
				replaceTextParagraph(p, values, replaceValues);
		} catch (Exception e) {
			log.error("Some error occured while trying to replace a value from a documents paragraph");
			log.error(e.getMessage());
		}			
	}
	
	/**
	 * Replaces a string value (or more) in the documents tables, by another String value
	 * 
	 * @param xwpfDoc
	 * @param value
	 * @param replaceValue
	 * @param onlyFirstOccurrence
	 */
	public static void replaceTextTables(XWPFDocument xwpfDoc, String[] values, String[] replaceValues) throws Exception {
		log.debug(TAG + ".replaceTextTables");

		if(values.length != replaceValues.length) {
			log.error("There are more values to replace than values to replace them with.");
			return;
		}		
		
		try {
			for (XWPFTable tbl : xwpfDoc.getTables())
				replaceTextTable(tbl, values, replaceValues);
		} catch (Exception e) {
			log.error("Some error occured while trying to replace a value from a documents table");
			log.error(e.getMessage());
		}
	}
	
	/**
	 * Replaces a string value (or more) in ONE table, by another String value
	 * 
	 * @param tbl
	 * @param values
	 * @param replaceValues
	 * @throws Exception
	 */
	private static void replaceTextTable(XWPFTable tbl, String[] values, String[] replaceValues) throws Exception {
		//log.debug(TAG + ".replaceTextTable");

		for (XWPFTableRow row : tbl.getRows())
			replaceTextTableRow(row, values, replaceValues);
	}

	/**
	 * Replaces a string value (or more) in ONE table ROW, by another String value
	 * 
	 * @param row
	 * @param values
	 * @param replaceValues
	 * @throws Exception
	 */
	private static void replaceTextTableRow(XWPFTableRow row, String[] values, String[] replaceValues) throws Exception {
		//log.debug(TAG + ".replaceTextTableRow");

		for (XWPFTableCell cell : row.getTableCells())
			replaceTextTableCell(cell, values, replaceValues);
	}
	
	/**
	 * Replaces a string value (or more) in ONE table CELL, by another String value

	 * @param cell
	 * @param values
	 * @param replaceValues
	 * @throws Exception
	 */
	private static void replaceTextTableCell(XWPFTableCell cell, String[] values, String[] replaceValues) throws Exception{
		//log.debug(TAG + ".replaceTextTableCell");

		for (XWPFParagraph p : cell.getParagraphs())
			replaceTextParagraph(p, values, replaceValues);
	}
	
	/**
	 * Replaces a string value (or more) in ONE paragraph, by another String value.
	 * Note a paragraph is divided in Runs. These runs can be like this (or any other form).
	 * Run01 - <<pl
	 * Run02 - aceholder
	 * Run03 - >> is here
	 * 
	 * @param p
	 * @param values
	 * @param replaceValues
	 * @throws Exception
	 */
	private static void replaceTextParagraph(XWPFParagraph p, String[] values, String[] replaceValues) throws Exception {
		//log.debug(TAG + ".replaceTextParagraph");

		List<XWPFRun> runs = p.getRuns();
		
		// Create full paragraph 
		if (runs != null) {
			try {
				// Create helperArray
				ReplaceHelper phh = new ReplaceHelper(runs);
				
				// Replace the placeholders
				replacePlaceholders(phh, values, replaceValues);
				
				// Update Runs
				updateRuns(runs, phh);
			} catch(Exception e) {
				log.error(e.getMessage());
				log.error("Could not read the paragraph");
			}
		}
	}
	
	/**
	 * Replace the values by other replaceValues, using the PlaceHolderHelper Object
	 * 
	 * @param phh
	 * @param values
	 * @param replaceValues
	 */
	private static void replacePlaceholders(ReplaceHelper phh, String[] values, String[] replaceValues) throws Exception {
		//log.debug(TAG + ".replacePlacholders");

		for(int i=0; i<values.length; i++)
			replacePlaceholder(phh, values[i], replaceValues[i]);
	}
	
	/**
	 * Replace a value by another replaceValue, using the PlaceHolderHelper Object
	 * 
	 * @param phh
	 * @param value
	 * @param replaceValue
	 */
	private static void replacePlaceholder(ReplaceHelper phh, String value, String replaceValue) throws Exception {
		//log.debug(TAG + ".replacePlacholder");
		int differenceLength = replaceValue.length() - value.length();

		if(phh.getParagraph().contains(value)) {
			// get offset to compare for position of placeholder
			int offset = phh.getParagraph().indexOf(value, 0);

			do {
				// replace 
				phh.setParagraph(phh.getParagraph().replaceFirst(value, replaceValue));
				
				if(differenceLength != 0)
					rearrangeHelperObject(phh, offset, differenceLength);
				
				offset = phh.getParagraph().indexOf(value, 0);
			} while(offset != -1);
		}
	}
	
	/**
	 * makes sure the document is cut up in nice runs when the runs are refilled.
	 * 
	 * @param phh
	 * @param offset
	 * @param differenceLength
	 */
	private static void rearrangeHelperObject(ReplaceHelper phh, int offset, int differenceLength) throws Exception {
		//log.debug(TAG + ".rearrangeHelperObject");
		boolean firstOccurrence = true;
		boolean ready = false;
		int runOffset = 0;

		// loop array to adjust length off runs
		for(int j=0; j<phh.length(); j++) {
			// build up the new offsets
			phh.setArrRuns(j, ReplaceHelper.RUN_OFFSET, runOffset);
			
			// is the offset of the value between the start and end of a run, 
			// rearrange the offsets and length of runs 
			// if ready is true, differenceLength is not negative anymore, 
			// so there is no need to shorten a run
			if(!ready && offset >= phh.getArrRuns(j, ReplaceHelper.RUN_OFFSET) && 
					offset < (phh.getArrRuns(j, ReplaceHelper.RUN_OFFSET) + phh.getArrRuns(j, ReplaceHelper.RUN_LENGTH))) {
				// new loop to start from the run the value is found an 
				// from this point change the length and offset from these runs
				for(int k=j; k<phh.length(); k++) {
					// check the character is the run, so they can remain after replacment
					int charInRun = (phh.getArrRuns(k, ReplaceHelper.RUN_OFFSET) + phh.getArrRuns(k, ReplaceHelper.RUN_LENGTH)) - offset;
					
					// set new length, make sure the part of the run is later in the right run
					if(!firstOccurrence || !(charInRun < -differenceLength)) {
						phh.setArrRuns(k, ReplaceHelper.RUN_LENGTH, phh.getArrRuns(k, ReplaceHelper.RUN_LENGTH) + differenceLength);
						firstOccurrence = false;
					}
					
					// if new length is negative, make it 0 and remember the rest
					// if the value is 0 or bigger, break the loop, set ready to true.
					if(phh.getArrRuns(k, ReplaceHelper.RUN_LENGTH) < 0) {
						differenceLength = phh.getArrRuns(k, ReplaceHelper.RUN_LENGTH);
						phh.setArrRuns(k, ReplaceHelper.RUN_LENGTH, 0);
					} else if(firstOccurrence) {
						firstOccurrence = false;
					} else {
						ready = true;
						break;
					}
				}
			}
			
			// set new offset, for next value
			runOffset = phh.getArrRuns(j, ReplaceHelper.RUN_OFFSET) + phh.getArrRuns(j, ReplaceHelper.RUN_LENGTH);
		}
	}
	
	/**
	 * Recreate the new runs
	 * 
	 * @param runs
	 * @param phh
	 */
	private static void updateRuns(List<XWPFRun> runs, ReplaceHelper phh) throws Exception {
		//log.debug(TAG + ".updateRuns");
		int offset=0;
		for (int i=0; i<runs.size(); i++) {
			//System.out.println(i + ": " + phh.getParagraph().substring(offset, offset + phh.getArrRuns(i, RUN_LENGTH)));

			if(phh.getArrRuns(i, ReplaceHelper.RUN_LENGTH) > 0)
				runs.get(i).setText(phh.getParagraph().substring(offset, offset + phh.getArrRuns(i, ReplaceHelper.RUN_LENGTH)), 0);
			else
				runs.get(i).setText("", 0);

			offset += phh.getArrRuns(i, ReplaceHelper.RUN_LENGTH);
		}		
	}
	
	
	/**
	 * Cannot remove runs getting a unmodifiable list error 
	 * 
	 * @param runs
	 * @param phh
	 */
	/*
	private static void removeEmptyRuns(List<XWPFRun> runs, PlaceHolderHelper phh) throws Exception {
		log.debug(TAG + ".removeEmptyRuns");

		//Remove empty Runs
		for (int i=runs.size(); i>0; --i) {
			if(phh.getArrRuns(i-1, RUN_LENGTH) == 0) {
				try {
					runs.remove(i-1);
				} catch(Exception e) {
					e.printStackTrace();
				}
			}
		}
	}
	*/
}