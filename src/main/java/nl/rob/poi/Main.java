package nl.rob.poi;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class Main {

	static String file = "c:\\temp\\a.docx";
	static String outfile = "c:\\temp\\a_out.docx";
	static HashMap<String, String> tags;

	public static void main(String[] args) throws FileNotFoundException, IOException {

		if (args.length > 0) {
			file = args[0];
			if (args.length > 1)
				outfile = args[1];
		}

		TagReplacer replacer = new TagReplacer();
		replacer.addTag("TAG", "tag");
		replacer.addTag("TAG2", "tag2");
		replacer.addTag("BOLD ITALIC", "bold italic");

		XWPFDocument doc = new XWPFDocument(
			new FileInputStream(file)
		);

		replacer.processBody(doc);

		FileOutputStream out = new FileOutputStream(outfile);
		doc.write(out);
		doc.close();
	}

}
