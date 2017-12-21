package br.com.assertsistemas.apachepoi.otherexamples;

import java.io.InputStream;
import java.net.ContentHandler;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.jar.Attributes;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.eclipse.persistence.internal.oxm.record.XMLReader;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

public class FromHowTo {
	public void processFirstSheet(String filename) throws Exception {
		OPCPackage pkg = OPCPackage.open(filename, PackageAccess.READ);
		try {
			XSSFReader r = new XSSFReader(pkg);
			SharedStringsTable sst = r.getSharedStringsTable();

			org.xml.sax.XMLReader parser = fetchSheetParser(sst);

			// process the first sheet
			InputStream sheet2 = r.getSheetsData().next();
			InputSource sheetSource = new InputSource(sheet2);
			parser.parse(sheetSource);
			sheet2.close();
		} finally {
			pkg.close();
		}
	}

	public void processAllSheets(String filename) throws Exception {
		OPCPackage pkg = OPCPackage.open(filename, PackageAccess.READ);
		try {
			XSSFReader r = new XSSFReader(pkg);
			SharedStringsTable sst = r.getSharedStringsTable();

			org.xml.sax.XMLReader parser = fetchSheetParser(sst);

			Iterator<InputStream> sheets = r.getSheetsData();
			while (sheets.hasNext()) {
				System.out.println("Processing new sheet:\n");
				InputStream sheet = sheets.next();
				InputSource sheetSource = new InputSource(sheet);
				parser.parse(sheetSource);
				sheet.close();
				System.out.println("");
			}
		} finally {
			pkg.close();
		}
	}

	public org.xml.sax.XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException {
		org.xml.sax.XMLReader parser = XMLReaderFactory.createXMLReader();
		org.xml.sax.ContentHandler handler = new SheetHandler(sst);
		parser.setContentHandler(handler);
		return parser;
	}

	/**
	 * See org.xml.sax.helpers.DefaultHandler javadocs
	 */
	private static class SheetHandler extends DefaultHandler {
		private final SharedStringsTable sst;
		private String lastContents;
		private boolean nextIsString;
		private boolean inlineStr;
		private final LruCache<Integer,String> lruCache = new LruCache<Integer,String>(50);

		private static class LruCache<A,B> extends LinkedHashMap<A, B> {
		    private final int maxEntries;

		    public LruCache(final int maxEntries) {
		        super(maxEntries + 1, 1.0f, true);
		        this.maxEntries = maxEntries;
		    }

		    @Override
		    protected boolean removeEldestEntry(final Map.Entry<A, B> eldest) {
		        return super.size() > maxEntries;
		    }
		}
		
		private SheetHandler(SharedStringsTable sst) {
			this.sst = sst;
		}

		@SuppressWarnings("unused")
		public void startElement(String uri, String localName, String name,
								 Attributes attributes) throws SAXException {
			// c => cell
			if(name.equals("c")) {
				// Print the cell reference
				System.out.print(attributes.getValue("r") + " - ");
				// Figure out if the value is an index in the SST
				String cellType = attributes.getValue("t");
				nextIsString = cellType != null && cellType.equals("s");
				inlineStr = cellType != null && cellType.equals("inlineStr");
			}
			// Clear contents cache
			lastContents = "";
		}

		@Override
        public void endElement(String uri, String localName, String name)
				throws SAXException {
			// Process the last contents as required.
			// Do now, as characters() may be called more than once
			if(nextIsString) {
				Integer idx = Integer.valueOf(lastContents);
				lastContents = lruCache.get(idx);
				if (lastContents == null && !lruCache.containsKey(idx)) {
				    lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
				    lruCache.put(idx, lastContents);
				}
				nextIsString = false;
			}

			// v => contents of a cell
			// Output after we've seen the string contents
			if(name.equals("v") || (inlineStr && name.equals("c"))) {
				System.out.println(lastContents);
			}
		}

		@Override
        public void characters(char[] ch, int start, int length) throws SAXException { // NOSONAR
			lastContents += new String(ch, start, length);
		}
	}

	public static void main(String[] args) throws Exception {
		FromHowTo howto = new FromHowTo();
		howto.processFirstSheet(args[0]);
		howto.processAllSheets(args[0]);
	}
}