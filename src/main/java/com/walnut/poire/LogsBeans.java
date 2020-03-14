package com.walnut.poire;

import java.io.File;
import java.io.IOException;
import java.util.Properties;

public class LogsBeans
{
  private static String maxFileSize = null;
  private static String logsFileName = null;
  private static String logsPattern = null;
  private static int maxBackupFiles = 1;
  private static String sFilePath = null;
  private static String sp = File.separator;
  private static Properties _props = Common.loadPropertyFile("PropertieFiles/logs.properties");
  
  protected static void setMaxFileSize(){
	  try{
		  maxFileSize = Common.getValueFromProperty(_props, "maxFileSize");
	  }catch (Exception e){
		  System.err.println(e.getMessage());
		  System.exit(0);
	  }
  }

  protected static String getMaxFileSize()
  {
    return maxFileSize;
  }
  
  protected static void setLogsFileName(){
	  try{
		  logsFileName = Common.getValueFromProperty(_props, "logsFileName");
	  }catch (Exception e){
		  System.err.println(e.getMessage());
		  System.exit(0);
	  }
  }
  
  protected static String getLogsFileName(){
	  return logsFileName;
  }
  
  protected static void setLogsPattern(){
	  try{
		  logsPattern = Common.getValueFromProperty(_props, "logsPattern");
	  }catch (Exception e){
		  System.err.println(e.getMessage());
		  System.exit(0);
	  }
  }
  
  protected static String getLogsPattern(){
	  return logsPattern;
  }

  protected static void setMaxBackupFiles()
  {
	  try{
		  maxBackupFiles = Integer.parseInt(
				  Common.getValueFromProperty(_props, "maxBackupFiles"));
	  }catch (Exception e){
		  System.err.println(e.getMessage());
		  System.exit(0);
	  }
  }
  
  protected static int getMaxBackupFiles(){
	  return maxBackupFiles;
  }
  
  protected static String setLogsFolder() throws IOException{
	  boolean result = false;
	  String sLogsFolderName = "logs";
	  String baseDirLoca = System.getProperty("user.dir");
	  File FolderPath = new File(baseDirLoca);
	  if(FolderPath.isDirectory()){
			File directory = new File(FolderPath.getCanonicalPath() + sp + sLogsFolderName);
			if (directory.isDirectory() && directory.exists()) {
				sFilePath = directory.getCanonicalPath();
			} else {
				result = directory.mkdirs();
				if(result){
					sFilePath = directory.getCanonicalPath();
				}
			}
		}else{
			System.err.println("Directory path is not found while creating a log folder");
			System.err.println(FolderPath.toString());
		}
	  return sFilePath;
  }
  
  protected static String getLogsFolderPath(){
	  return sFilePath;
  }
  
}
