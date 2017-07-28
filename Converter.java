import java.awt.Color;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Iterator;

import javax.imageio.ImageIO;
import javax.imageio.ImageWriter;
import javax.imageio.stream.ImageOutputStream;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;

import org.apache.poi.hssf.converter.ExcelToHtmlConverter;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.docx4j.Docx4J;
import org.docx4j.convert.out.FOSettings;
import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFont;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.model.fields.FieldUpdater;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.w3c.dom.Document;
import org.xhtmlrenderer.pdf.ITextRenderer;

import com.lowagie.text.DocumentException;
import com.lowagie.text.Image;
import com.lowagie.text.Rectangle;
import com.lowagie.text.pdf.BaseFont;
import com.lowagie.text.pdf.PdfPCell;
import com.lowagie.text.pdf.PdfPTable;
import com.lowagie.text.pdf.PdfWriter;

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
 * 
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
	 * @param inPath
	 *            The input file path
	 * @param outPath
	 *            The output file path. If path format is not pdf, will be
	 *            changed to pdf. Put null to generate pdf file in the same
	 *            directory with the same name
	 * @param mainFontUsed
	 * 			  The main font used in the document. Required for font mapping. 
	 * 			  Note: Font mapping for Arial, Times New Roman, Calibri, Helvetica, 等线 is included
	 * @throws Exception 
	 */
	public static void docxToPDF(String inPath, String outPath, String mainFontUsed) throws Exception {
		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new File(inPath));
		
		//TODO: Support Chinese font
		//Set up font mapper
		String os = System.getProperty("os.name");
		String regex;
		if(os.toLowerCase().startsWith("windows")){
			regex =".*(calibri|camb|cour|arial|times|comic|georgia|impact|LSANS|pala|tahoma|trebuc|verdana|symbol|webdings|wingding).*";
		}else{
			regex =".*(Courier New|Arial|Times New Roman|Comic Sans|Georgia|Impact|Lucida Console|Lucida Sans Unicode|Palatino Linotype|Tahoma|Trebuchet|Verdana|Symbol|Webdings|Wingdings|MS Sans Serif|MS Serif|).*";
		}	
		PhysicalFonts.setRegex(regex);
	    Mapper fontMapper = new IdentityPlusMapper();
	    wordMLPackage.setFontMapper(fontMapper);
	    PhysicalFonts.discoverPhysicalFonts();
	    PhysicalFont font = PhysicalFonts.get("Arial Unicode MS");
	    System.out.println(font);
	    if (font!=null) {
	    	fontMapper.put("Times New Roman", font);
	        fontMapper.put("Arial", font);
	        fontMapper.put("Calibri", font);
	        fontMapper.put("Helvetica", font);
	        fontMapper.put("等线", font);
	        fontMapper.put("宋体", font);
	        if(mainFontUsed != null){
	        	fontMapper.put(mainFontUsed, font);
	        }
	    }
	    fontMapper.put("Libian SC Regular", PhysicalFonts.get("SimSun"));
	   
		// Refresh the values of DOCPROPERTY fields
		FieldUpdater updater = new FieldUpdater(wordMLPackage);
		updater.update(true);

		// Validate outPath
		outPath = pathValidator(inPath, outPath);

		// Setup
		FOSettings foSettings = Docx4J.createFOSettings();
		foSettings.setWmlPackage(wordMLPackage);

		// Output
		OutputStream out = new FileOutputStream(outPath);
		Docx4J.toFO(foSettings, out, Docx4J.FLAG_EXPORT_PREFER_XSL);
		System.out.println("Saved: " + outPath);

		// Cleanup
		out.flush();
		out.close();
		if (wordMLPackage.getMainDocumentPart().getFontTablePart() != null) {
			wordMLPackage.getMainDocumentPart().getFontTablePart().deleteEmbeddedFontTempFiles();
		}
		updater = null;
		wordMLPackage = null;
	}

	/**
	 * Converts xls to pdf. Doesn't support chart conversion (chart will not
	 * show up in pdf)
	 * 
	 * @param inPath
	 *            The input file path
	 * @param outPath
	 *            The output file path. If path format is not pdf, will be
	 *            changed to pdf. Put null to generate pdf file in the same
	 *            directory with the same name
	 * @throws ParserConfigurationException
	 * @throws IOException
	 * @throws DocumentException
	 * @throws Exception
	 */
	public static void xlsToPDF(String inPath, String outPath)
			throws IOException, ParserConfigurationException, DocumentException {
		// Convert input file into HTML
		Document inHTML = ExcelToHtmlConverter.process(new File(inPath));

		// Validate outPath
		outPath = pathValidator(inPath, outPath);

		// Convert to PDF
		htmlToPDF(inHTML, outPath);
	}

	/**
	 * Converts xlsx to pdf. Chart and color formatting conversion not supported
	 * 
	 * @param inPath
	 *            The input file path
	 * @param outPath
	 *            The output file path. If path format is not pdf, will be
	 *            changed to pdf. Put null to generate pdf file in the same
	 *            directory with the same name
	 * @throws IOException
	 * @throws DocumentException
	 * @throws ParserConfigurationException
	 * @throws TransformerException 
	 */
	public static void xlsxToPDF(String inPath, String outPath, boolean outputColumnHeader, boolean outputRowNumber) throws IOException, DocumentException, ParserConfigurationException, TransformerException {
		//Load file
		FileInputStream fis = new FileInputStream(new File(inPath));
		
		// Convert input file into HTML
		Document inHTML = XLSXToHTMLConverter.convert(new XSSFWorkbook(fis), outputColumnHeader, outputRowNumber);
		
		fis.close();
		
		// Validate outPath
		outPath = pathValidator(inPath, outPath);
		
		//TODO: HTML output for font debugging
//		TransformerFactory tranFactory = TransformerFactory.newInstance();
//		Transformer aTransformer = tranFactory.newTransformer();
//		Source src = new DOMSource(inHTML);
//		Result dest = new StreamResult(new File(outPath + ".html"));
//		aTransformer.transform(src, dest);

		// Convert to PDF
		htmlToPDF(inHTML, outPath);
	}

	/**
	 * Converts pptx to PDF file 
	 * 
	 * @param inPath The input file path
	 * @param outPath
	 * 			  The output file path. If path format is not pdf, will be
	 *            changed to pdf. Put null to generate pdf file in the same
	 *            directory with the same name
	 * @throws Exception
	 */
	public static void pptxToPDF(String inPath, String outPath) throws Exception {
        
		//Load file
		FileInputStream fis = new FileInputStream(new File(inPath));
		XMLSlideShow inPPT = new XMLSlideShow(fis);
		byte[] byteImgData;
		fis.close();
		
		// Validate outPath
		outPath = pathValidator(inPath, outPath);
		
		//Dimesions
		Dimension pgsize = inPPT.getPageSize();
		int width = (int)pgsize.width;
		int height = (int)pgsize.height;
		
		//Setup document
		com.lowagie.text.Document document = new com.lowagie.text.Document();
		PdfWriter.getInstance(document, new FileOutputStream(outPath));
		PdfPTable table = new PdfPTable(1);	
		
		//Convert each slide into image
		int i = 0;
		for(XSLFSlide slide : inPPT.getSlides()){
			
			BufferedImage slideImg = new BufferedImage(pgsize.width, pgsize.height, BufferedImage.TYPE_INT_RGB);
			
			//G2D setup
			Graphics2D g2d = slideImg.createGraphics();
			g2d.setPaint(Color.white);
			g2d.fill(new Rectangle2D.Float(0, 0, pgsize.width, pgsize.height));
			g2d.clearRect(0, 0, width, height);
			
			slide.draw(g2d);
			
			//Save image into byte array
			System.out.println("Drawing slide: " + (i+1));
			ByteArrayOutputStream slideDrawn = new ByteArrayOutputStream();
			ImageOutputStream outputStream = ImageIO.createImageOutputStream(slideDrawn);
			Iterator<ImageWriter> iterator = ImageIO.getImageWritersByFormatName("png");
			if(!iterator.hasNext()){
				throw new IllegalStateException("Writers Not Found");
			}
			ImageWriter imageWriter = iterator.next();
			imageWriter.setOutput(outputStream);
			imageWriter.write(slideImg);
			byteImgData = slideDrawn.toByteArray();
			
			//Printing to PDF
			System.out.println("Printing slide: " + (i+1));			
			Image img = Image.getInstance(byteImgData);
			document.setPageSize(new Rectangle(img.getWidth(), img.getHeight()));
			document.open();
			img.setAbsolutePosition(0, 0);
			table.addCell(new PdfPCell(img, true));
			i++;
		}
		document.add(table);
		inPPT.close();
	    document.close();
	    System.out.println("Saved: " + outPath);
	}

	/**
	 * Private method to convert HTML in w3c dom document to PDF using flyingSaucer
	 * 
	 * @param in
	 *            Input document
	 * @param outPath
	 *            Output path
	 * @throws DocumentException
	 * @throws IOException
	 */
	private static void htmlToPDF(Document in, String outPath) throws DocumentException, IOException {
		ITextRenderer renderer = new ITextRenderer();
		renderer.setDocument(in, null);
		renderer.layout();
		
		//Set font
		renderer.getFontResolver().addFont("fonts/ARIALUNI.TTF", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);

		//Output
		OutputStream os = new FileOutputStream(outPath);
		renderer.createPDF(os);
		
		System.out.println("Saved: " + outPath);

		// Cleanup
		os.flush();
		os.close();
		renderer = null;	

	}

	/**
	 * Check if outPath contains ".pdf"
	 * 
	 * @param inPath
	 *            Input file path
	 * @param outPath
	 *            Output file path
	 * @return
	 */
	private static String pathValidator(String inPath, String outPath) {
		if (outPath == null) {
			return inPath.substring(0, inPath.indexOf('.')) + ".pdf";
		} else if (!outPath.contains(".")) {
			return outPath += ".pdf";
		} else if (!outPath.endsWith("pdf")) {																			
			return outPath.substring(0, outPath.indexOf('.')) + ".pdf";
		} else if (outPath.endsWith("pdf")) {
			return outPath;
		} else {
			return inPath.substring(0, inPath.indexOf('.')) + ".pdf";
		}
	}
}
