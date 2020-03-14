package com.walnut.poire.slides;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import poire.office.framework.Common;
import org.apache.poi.hslf.model.Slide;
import org.apache.poi.hslf.usermodel.SlideShow;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

public class PoirePowerPointSlide {

	/**
	 * @author Ssire(siri) kumar Puttagunta</br>
	 * @version 2.0.0</br>
	 *          <p>
	 *          The below class methods are Designed with the help of POI jars.
	 *          These methods are useful for doing operations on the Power point
	 *          File formats. The whole class methods are in static only and
	 *          these method are useful in all the word Power point formats like
	 *          .ppt and .pptx
	 *          </p>
	 */

	private static boolean pptx = false;

	/**
	 * It will used to create the Power Point file in given File Format
	 * @param  pptxfileFormat  - Boolean value for .pptx format true for .pptx file.
	 * @param  filePath  - Directory location path for to create the excel file.     
	 * @param  fileName  - Name of the file for excel file in the Directory. 
	 *  */
	public static File createPowerPointFile(boolean pptxfileFormat, String filePath, String fileName){
		FileOutputStream fileOut = null;
		File _sReturnFileObj = null;
		File _sFilePathObj = null;
		SlideShow _ppt = null;
		XMLSlideShow _pptx = null;
		XSLFSlide slidx = null;
		Slide slid = null;
		try {
			_sFilePathObj = new File(filePath);
			pptx = pptxfileFormat;
			System.out.println("Creating the Power Point file");
			if(pptx){
				_sFilePathObj = new File(filePath + "\\"+ fileName + ".pptx");
				_pptx = new XMLSlideShow();
				slidx =  _pptx.createSlide();
				slidx.getTitle();
				fileOut = new FileOutputStream(_sFilePathObj);
				_pptx.write(fileOut);
				fileOut.close();
				_sReturnFileObj = _sFilePathObj.getAbsoluteFile();
				System.out.println("Created " + filePath + "\\"+ fileName + ".pptx" + " Word Document file");
			}else{
				_sFilePathObj = new File(filePath + "\\"+ fileName + ".ppt");
				_ppt = new SlideShow();
				slid =  _ppt.createSlide();
				slid.getTitle();
				fileOut = new FileOutputStream(_sFilePathObj);
				_ppt.write(fileOut);
				fileOut.close();
				_sReturnFileObj = _sFilePathObj.getAbsoluteFile();
				System.out.println("Created " + filePath + "\\"+ fileName + ".ppt" + " Power Point file");
			}
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
	 * powerPointFileType(File powerPointFile)
	 * 
	 * */
	private static String powerPointFileType(File powerPointFile){
		String sReturnFileType = "";
		try {
			if(powerPointFile.isFile() && powerPointFile.exists()){
				String fileName = powerPointFile.getName();
				String extension = fileName.substring(fileName.lastIndexOf(".") + 1, fileName.length());
				if (extension.equalsIgnoreCase("ppt")){
					sReturnFileType = "ppt";
				}else if(extension.equalsIgnoreCase("pptx")){
					sReturnFileType = "pptx";
				}else{
					System.err.println("The File is not a type of Power Point presentation file format");
				} 
			}	
		} catch (Exception e) {
			Common.addExceptionLogger(e);
		}
		return sReturnFileType;
	}
	
	
	
	
	
	
	
	
	
	
	
}
