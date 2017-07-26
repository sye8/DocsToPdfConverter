import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import javax.xml.parsers.ParserConfigurationException;

import org.apache.poi.hssf.converter.ExcelToHtmlConverter;
import org.docx4j.Docx4J;
import org.docx4j.convert.out.FOSettings;
import org.docx4j.model.fields.FieldUpdater;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.w3c.dom.Document;
import org.xhtmlrenderer.pdf.ITextRenderer;

import com.lowagie.text.DocumentException;


/*
 * MIT License
 * 
 * Copyright (c) 2017 Sifan YE
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
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

/**
 * Uses Docs4J to convert documents to pdf.
 * 
 * For docx conversion, docx4j-export-FO is needed
 * @see https://github.com/plutext/docx4j-export-FO
 * 
 * 
 * 
 * @author yesifan
 *
 * 
 */
public class Converter {

	/**
	 * Converts .docx files to pdf
	 * 
	 * @param inPath The input file path
	 * @param outPath The output file path. If path format is not pdf, will be changed to pdf. Put null to generate pdf file in the same directory with the same name
	 * @throws Docx4JException
	 * @throws IOException
	 */
	public static void docxToPDF(String inPath, String outPath) throws Docx4JException, IOException{
		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new File(inPath));
		
		// Refresh the values of DOCPROPERTY fields 
		FieldUpdater updater = new FieldUpdater(wordMLPackage);
		updater.update(true);
		
		//Validate outPath
		outPath = pathValidator(inPath, outPath);
		
		//Setup
		FOSettings foSettings = Docx4J.createFOSettings();
		foSettings.setWmlPackage(wordMLPackage);
		
		//Output
		OutputStream out = new FileOutputStream(outPath);
		Docx4J.toFO(foSettings, out, Docx4J.FLAG_EXPORT_PREFER_XSL);
		System.out.println("Saved: " + outPath);
		
		
		//Cleanup
		out.flush();
		out.close();
		if (wordMLPackage.getMainDocumentPart().getFontTablePart()!=null) {
			wordMLPackage.getMainDocumentPart().getFontTablePart().deleteEmbeddedFontTempFiles();
		}		
		updater = null;
		wordMLPackage = null;
	}
	
	/**
	 * Converts xls to pdf. Doesn't support chart conversion (chart will not show up in pdf)
	 * 
	 * @param inPath The input file path
	 * @param outPath The output file path. If path format is not pdf, will be changed to pdf. Put null to generate pdf file in the same directory with the same name
	 * @throws ParserConfigurationException 
	 * @throws IOException 
	 * @throws DocumentException 
	 * @throws Exception
	 */
	public static void xlsToPDF(String inPath, String outPath) throws IOException, ParserConfigurationException, DocumentException{
		//Convert input file into HTML
		Document inHTML = ExcelToHtmlConverter.process(new File(inPath));
		
		ITextRenderer renderer = new ITextRenderer();
		renderer.setDocument(inHTML, null);
		
		renderer.layout();
		
		//Validate outPath
		outPath = pathValidator(inPath, outPath);
		
		OutputStream os = new FileOutputStream(outPath);
		
		renderer.createPDF(os);
		System.out.println("Saved: " + outPath);
		
		//Cleanup
		os.flush();
		os.close();
		renderer = null;
	}
	
	/**
	 * Check if outPath contains ".pdf"
	 * 
	 * @param inPath Input file path
	 * @param outPath Output file path
	 * @return
	 */
	private static String pathValidator(String inPath, String outPath){
		if(outPath == null){
			return inPath.substring(0, inPath.indexOf('.')) + ".pdf";
	 	}else if(!outPath.contains(".")){
			return outPath += ".pdf";
		}else if(!outPath.substring(outPath.indexOf('.')+1).equals("pdf")){ //Check if outPath contains valid file type
			return outPath.substring(0, outPath.indexOf('.')) + ".pdf";
		}else{
			return inPath.substring(0, inPath.indexOf('.')) + ".pdf";
		}
	}
}
