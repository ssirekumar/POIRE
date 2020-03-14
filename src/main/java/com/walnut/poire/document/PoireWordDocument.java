package com.walnut.poire.document;

import com.walnut.poire.Common;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;
import java.util.Arrays;
import org.apache.poi.hssf.record.BottomMarginRecord;
import org.apache.poi.hssf.record.LeftMarginRecord;
import org.apache.poi.hssf.record.RecordInputStream;
import org.apache.poi.hssf.record.RightMarginRecord;
import org.apache.poi.hssf.record.TopMarginRecord;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;



public class PoireWordDocument{
	
	/**
	 * @author Ssire kumar Puttagunta</br>
	 * @version 2.0.0</br>
	 *          <p>
	 *          The below class methods are Designed with the help of POI jars.
	 *          These methods are useful for doing operations on the word
	 *          Document File formats. The whole class methods are in static
	 *          only and these method are useful in all the word file formats
	 *          like .doc and .docx
	 *          </p>
	 */

	private static boolean docx = false;

	 /**
	 * It will used to set the Word Document Margins with the help of the IWordMargins interface enum values.
	 * @param  filePath  - Path of the file which is .docx or .doc     
	 * @param  marginType  - type of margin Exe: IWordMargins.NARROW.values().
	 *                       Use the IWordMargins interface of enums. 
	 *         Exe: IWordMargins.NARROW.values(), IWordMargins.NORMAL.values(), IWordMargins.MODERATE.values()  
	 *              IWordMargins.WIDE.values(), IWordMargins.OFFICE_2003_DEFAULT.values(),           
	 *  */
	public static void setMarginsOfdoc(String filePath, Object[] marginType){     //Incomplete method at .doc file margin setting.
		File sFilePathObj = null;
		OPCPackage _OPCXWPF = null;
		POIFSFileSystem  _FSHWPF_SSIRE = null;
		InputStream _iStreem = null;
		FileOutputStream fileOut = null;
		boolean bStreenClose = false;
		XWPFDocument _docx = null;
		HWPFDocument _doc = null;
		try {
			sFilePathObj = new File(filePath);
			docx = PoireWordDocument.wordFileType(sFilePathObj).equalsIgnoreCase("docx");
			if(docx){
				_OPCXWPF = OPCPackage.open(new FileInputStream(sFilePathObj));
				_docx = new XWPFDocument(_OPCXWPF);
				CTSectPr sectPr = _docx.getDocument().getBody().addNewSectPr();
				CTPageMar pageMar = sectPr.addNewPgMar();
				if(Arrays.equals(marginType, IWordMargins.NORMAL.values())){
					pageMar.setBottom(BigInteger.valueOf(Long.parseLong(String.valueOf(IWordMargins.NORMAL.BottomMargin.getValue()))));
					pageMar.setTop(BigInteger.valueOf(Long.parseLong(String.valueOf(IWordMargins.NORMAL.TopMargin.getValue()))));
					pageMar.setLeft(BigInteger.valueOf(Long.parseLong(String.valueOf(IWordMargins.NORMAL.LeftMargin.getValue()))));
					pageMar.setRight(BigInteger.valueOf(Long.parseLong(String.valueOf(IWordMargins.NORMAL.RightMargin.getValue()))));
				}else if(Arrays.equals(marginType, IWordMargins.NARROW.values())){
					pageMar.setBottom(BigInteger.valueOf(Long.parseLong(String.valueOf(IWordMargins.NARROW.BottomMargin.getValue()))));
					pageMar.setTop(BigInteger.valueOf(Long.parseLong(String.valueOf(IWordMargins.NARROW.TopMargin.getValue()))));
					pageMar.setLeft(BigInteger.valueOf(Long.parseLong(String.valueOf(IWordMargins.NARROW.LeftMargin.getValue()))));
					pageMar.setRight(BigInteger.valueOf(Long.parseLong(String.valueOf(IWordMargins.NARROW.RightMargin.getValue()))));
				}else if(Arrays.equals(marginType, IWordMargins.MODERATE.values())){
					pageMar.setBottom(BigInteger.valueOf(Long.parseLong(String.valueOf(IWordMargins.MODERATE.BottomMargin.getValue()))));
					pageMar.setTop(BigInteger.valueOf(Long.parseLong(String.valueOf(IWordMargins.MODERATE.TopMargin.getValue()))));
					pageMar.setLeft(BigInteger.valueOf(Long.parseLong(String.valueOf(IWordMargins.MODERATE.LeftMargin.getValue()))));
					pageMar.setRight(BigInteger.valueOf(Long.parseLong(String.valueOf(IWordMargins.MODERATE.RightMargin.getValue()))));
				}else if (Arrays.equals(marginType, IWordMargins.WIDE.values())){
					pageMar.setBottom(BigInteger.valueOf(Long.parseLong(String.valueOf(IWordMargins.WIDE.BottomMargin.getValue()))));
					pageMar.setTop(BigInteger.valueOf(Long.parseLong(String.valueOf(IWordMargins.WIDE.TopMargin.getValue()))));
					pageMar.setLeft(BigInteger.valueOf(Long.parseLong(String.valueOf(IWordMargins.WIDE.LeftMargin.getValue()))));
					pageMar.setRight(BigInteger.valueOf(Long.parseLong(String.valueOf(IWordMargins.WIDE.RightMargin.getValue()))));
				}else if(Arrays.equals(marginType, IWordMargins.OFFICE_2003_DEFAULT.values())){
					pageMar.setBottom(BigInteger.valueOf(Long.parseLong(String.valueOf(IWordMargins.OFFICE_2003_DEFAULT.BottomMargin.getValue()))));
					pageMar.setTop(BigInteger.valueOf(Long.parseLong(String.valueOf(IWordMargins.OFFICE_2003_DEFAULT.TopMargin.getValue()))));
					pageMar.setLeft(BigInteger.valueOf(Long.parseLong(String.valueOf(IWordMargins.OFFICE_2003_DEFAULT.LeftMargin.getValue()))));
					pageMar.setRight(BigInteger.valueOf(Long.parseLong(String.valueOf(IWordMargins.OFFICE_2003_DEFAULT.RightMargin.getValue()))));
				}else {
					System.err.println("Margin type is not defined as per the IWordMargins interface");
				}
				fileOut = new FileOutputStream(sFilePathObj);
				_docx.write(fileOut);
				fileOut.flush();
				fileOut.close();
				bStreenClose = true;
			}else{
				_iStreem = new FileInputStream(sFilePathObj);
				RecordInputStream r = new RecordInputStream(_iStreem);
				_iStreem = new FileInputStream(sFilePathObj);
				_FSHWPF_SSIRE = new POIFSFileSystem(_iStreem);
			 	_doc = new HWPFDocument(_FSHWPF_SSIRE);
			 	BottomMarginRecord bm = new BottomMarginRecord(r);
                TopMarginRecord tm = new TopMarginRecord(r);
                RightMarginRecord rm = new RightMarginRecord(r);
                LeftMarginRecord lm = new LeftMarginRecord(r);
                if(Arrays.equals(marginType, IWordMargins.NORMAL.values())){
                	 bm.setMargin((double)Long.parseLong(String.valueOf(IWordMargins.NORMAL.BottomMargin.getValue())));
                     tm.setMargin((double)Long.parseLong(String.valueOf(IWordMargins.NORMAL.TopMargin.getValue())));
                     lm.setMargin((double)Long.parseLong(String.valueOf(IWordMargins.NORMAL.LeftMargin.getValue())));
                     rm.setMargin((double)Long.parseLong(String.valueOf(IWordMargins.NORMAL.RightMargin.getValue())));
                }else if(Arrays.equals(marginType, IWordMargins.NARROW.values())){
                	 bm.setMargin((double)Long.parseLong(String.valueOf(IWordMargins.NARROW.BottomMargin.getValue())));
                     tm.setMargin((double)Long.parseLong(String.valueOf(IWordMargins.NARROW.TopMargin.getValue())));
                     lm.setMargin((double)Long.parseLong(String.valueOf(IWordMargins.NARROW.LeftMargin.getValue())));
                     rm.setMargin((double)Long.parseLong(String.valueOf(IWordMargins.NARROW.RightMargin.getValue())));
                }else if (Arrays.equals(marginType, IWordMargins.MODERATE.values())) {
                	 bm.setMargin((double)Long.parseLong(String.valueOf(IWordMargins.MODERATE.BottomMargin.getValue())));
                     tm.setMargin((double)Long.parseLong(String.valueOf(IWordMargins.MODERATE.TopMargin.getValue())));
                     lm.setMargin((double)Long.parseLong(String.valueOf(IWordMargins.MODERATE.LeftMargin.getValue())));
                     rm.setMargin((double)Long.parseLong(String.valueOf(IWordMargins.MODERATE.RightMargin.getValue())));
				}else if(Arrays.equals(marginType, IWordMargins.WIDE.values())){
					 bm.setMargin((double)Long.parseLong(String.valueOf(IWordMargins.WIDE.BottomMargin.getValue())));
                     tm.setMargin((double)Long.parseLong(String.valueOf(IWordMargins.WIDE.TopMargin.getValue())));
                     lm.setMargin((double)Long.parseLong(String.valueOf(IWordMargins.WIDE.LeftMargin.getValue())));
                     rm.setMargin((double)Long.parseLong(String.valueOf(IWordMargins.WIDE.RightMargin.getValue())));
				}else if(Arrays.equals(marginType, IWordMargins.OFFICE_2003_DEFAULT.values())){
					 bm.setMargin((double)Long.parseLong(String.valueOf(IWordMargins.OFFICE_2003_DEFAULT.BottomMargin.getValue())));
                     tm.setMargin((double)Long.parseLong(String.valueOf(IWordMargins.OFFICE_2003_DEFAULT.TopMargin.getValue())));
                     lm.setMargin((double)Long.parseLong(String.valueOf(IWordMargins.OFFICE_2003_DEFAULT.LeftMargin.getValue())));
                     rm.setMargin((double)Long.parseLong(String.valueOf(IWordMargins.OFFICE_2003_DEFAULT.RightMargin.getValue())));
				}else{
					
				}
                fileOut = new FileOutputStream(sFilePathObj);
				_doc.write(fileOut);
				fileOut.flush();
				fileOut.close();
				bStreenClose = true;
			}
			if(docx){
				_OPCXWPF.close();
				bStreenClose = true;
			}else{
				_iStreem.close();
				bStreenClose = true;
			}
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(docx){
				try {
					if(!bStreenClose){
						_OPCXWPF.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_iStreem.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}
	
	/**
	 * It will used to create the .docx Document file.
	 * @param  filePath  - Directory location path for to create the excel file.     
	 * @param  fileName  - Name of the file for excel file in the Directory. 
	 *  */
	public static File createdocxFile(String filePath, String fileName){
		FileOutputStream fileOut = null;
		File _sReturnFileObj = null;
		File _sFilePathObj = null;
		XWPFDocument _docx = null;
		try {
			System.out.println("Creating the Word Document file");
			_sFilePathObj = new File(filePath);
			_sFilePathObj = new File(filePath + "\\"+ fileName + ".docx");
			_docx = new XWPFDocument();
			fileOut = new FileOutputStream(_sFilePathObj);
			_docx.write(fileOut);
			fileOut.flush();
			fileOut.close();
			_sReturnFileObj = _sFilePathObj.getAbsoluteFile();
			System.out.println("Created " + filePath + "\\"+ fileName + ".docx" + " Word Document file");
		} catch (IOException Io) {
			Common.addExceptionLogger(Io);
		} catch (Exception e){
			Common.addExceptionLogger(e);
		}finally {
			if (fileOut != null){
				try {
					fileOut.close();
				} catch (IOException e) {
					e.printStackTrace();
					Common.addExceptionLogger(e);
				}
			}
		}
		return _sReturnFileObj;
	}
		
   	/**
	 * It will used to create the Word Document file in given File Format
	 * @param  xlsxfileFormat  - Boolean value for .docx format true for .docx file.
	 * @param  filePath  - Directory location path for to create the excel file.     
	 * @param  fileName  - Name of the file for excel file in the Directory. 
	 * */
	public static File createDocumentFile(boolean docxfileFormat, String filePath, String fileName){
		FileOutputStream fileOut = null;
		File _sReturnFileObj = null;
		File _sFilePathObj = null;
		XWPFDocument _docx = null;
		HWPFDocument _doc = null;
		InputStream _iStreem = null;
		try {
			_sFilePathObj = new File(filePath);
			docx = docxfileFormat;
			if(docx){
				_sFilePathObj = new File(filePath + "\\"+ fileName + ".docx");
				_docx = new XWPFDocument();
				fileOut = new FileOutputStream(_sFilePathObj);
				_docx.write(fileOut);
				fileOut.close();
				_sReturnFileObj = _sFilePathObj.getAbsoluteFile();
			}else{
				_sFilePathObj = new File(filePath + "\\"+ fileName + ".doc");
				_iStreem = ClassLoader.getSystemResourceAsStream("office/word/DocTemplate.doc");
				_doc = new HWPFDocument(_iStreem);
				_iStreem.close(); // This stream is no longer needed after doc is created
				fileOut = new FileOutputStream(_sFilePathObj);
				_doc.write(fileOut);
				fileOut.close();
				_sReturnFileObj = _sFilePathObj.getAbsoluteFile();
			}
			System.out.println("Creating the Word Document file");
		} catch (IOException Io) {
			Common.addExceptionLogger(Io);
		} catch (Exception e){
			Common.addExceptionLogger(e);
		}finally {
			if (fileOut != null){
				try {
					fileOut.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return _sReturnFileObj;
	}
    
	/**
	 * It will used to create the Document paragraph data default font family as 'Times New Roman'
	 * and font size as 12, ParagraphAlignment.LEFT, WordWrap true.
	 * @param  filePath  - Path of the file which is .docx or .doc     
	 * @param  paragraphText  - Paragraph text need to write in the document.           
	 *  */
	public static void createParagraph(String filePath, String paragraphText){  //incomplete method at while printing continus paragraphs.
		File sFilePathObj = null;
		OPCPackage _OPCXWPF = null;
		POIFSFileSystem  _FSHWPF_SSIRE = null;
		InputStream _iStreem = null;
		FileOutputStream fileOut = null;
		boolean bStreenClose = false;
		XWPFDocument _docx = null;
		HWPFDocument _doc = null;
		try {
			sFilePathObj = new File(filePath);
			docx = PoireWordDocument.wordFileType(sFilePathObj).equalsIgnoreCase("docx");
			if(docx){
				_OPCXWPF = OPCPackage.open(new FileInputStream(sFilePathObj));
				_docx = new XWPFDocument(_OPCXWPF);
				XWPFParagraph paragraph = _docx.createParagraph();
				paragraph.setWordWrap(true);
				paragraph.setAlignment(ParagraphAlignment.LEFT);
				XWPFRun run = paragraph.createRun();
				run.setFontFamily("Times New Roman");
				run.setFontSize(12);
				run.setText(paragraphText);
				fileOut = new FileOutputStream(sFilePathObj);
				_docx.write(fileOut);
				fileOut.flush();
				fileOut.close();
				bStreenClose = true;
			}else{
				_iStreem = new FileInputStream(sFilePathObj);
				_FSHWPF_SSIRE = new POIFSFileSystem(_iStreem);
			 	_doc = new HWPFDocument(_FSHWPF_SSIRE);
			 	Range range = _doc.getRange();
			 	Paragraph paragraph = range.getParagraph(0);
			 	paragraph.setWordWrapped(true);
			 	paragraph.setFontAlignment(0);
			 	CharacterRun _cRun = paragraph.insertBefore(paragraphText + "\n");
			 	_cRun.setFontSize(12*2);
			 	fileOut = new FileOutputStream(sFilePathObj);
				_doc.write(fileOut);
				fileOut.flush();
				fileOut.close();
				bStreenClose = true;
			}
			if(docx){
				_OPCXWPF.close();
				bStreenClose = true;
			}else{
				_iStreem.close();
				bStreenClose = true;
			}
		}catch(FileNotFoundException fe){
			Common.addExceptionLogger(fe);
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}finally{
			if(docx){
				try {
					if(!bStreenClose){
						_OPCXWPF.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				} 
			}else{
				try {
					if(!bStreenClose){
						_iStreem.close();
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}
		
	/**
	 * wordFileType(File excelFile)
	 * 
	 * */
	private static String wordFileType(File excelFile){
		String sReturnFileType = "";
		try {
			if(excelFile.isFile() && excelFile.exists()){
				String fileName = excelFile.getName();
				String extension = fileName.substring(fileName.lastIndexOf(".") + 1, fileName.length());
				if (extension.equalsIgnoreCase("doc")){
					sReturnFileType = "doc";
				}else if(extension.equalsIgnoreCase("docx")){
					sReturnFileType = "docx";
				}else{
					System.err.println("The File is not a type of Word Document file format");
				} 
			}	
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}
		return sReturnFileType;
	}

    












}
