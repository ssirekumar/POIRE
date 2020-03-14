package com.walnut.poire;

import java.io.File;

import org.apache.log4j.Appender;
import org.apache.log4j.AppenderSkeleton;
import org.apache.log4j.Level;
import org.apache.log4j.Logger;
import org.apache.log4j.PatternLayout;
import org.apache.log4j.RollingFileAppender;
import org.apache.log4j.spi.LoggingEvent;

/**
 * This class is to create the execution logs for automation framework
 * @author Siri Kumar Puttagunta*/
public class LoggerPoire
{
	private static boolean isLoggingInitialized = false;
	private static RollingFileAppender _rfa = null;
	private static Logger _log = null;
	private static String sp = File.separator;

	private static void setLoggerConfiguration() {
		try {
			isLoggingInitialized = true;
			LogsBeans.setMaxFileSize();
			LogsBeans.setLogsFileName();
			LogsBeans.setLogsPattern();
			LogsBeans.setMaxBackupFiles();
			LogsBeans.setLogsFolder();
			_rfa = new RollingFileAppender();
			_rfa.setMaxFileSize(LogsBeans.getMaxFileSize());
			_rfa.setMaxBackupIndex(LogsBeans.getMaxBackupFiles());
			_rfa.setName("Poire-Logger");
			_rfa.setFile(LogsBeans.getLogsFolderPath() + sp
					+ LogsBeans.getLogsFileName());
			_rfa.setLayout(new PatternLayout(LogsBeans.getLogsPattern()));
			_rfa.setImmediateFlush(true);
			_rfa.setThreshold(Level.ALL);
			_rfa.setAppend(true);
			_rfa.activateOptions();
			Logger.getRootLogger().addAppender(_rfa);
		} catch (Exception e) {
			System.err.println(e.getMessage());
			System.exit(0);
		}
	}

	/**
	 * Function to set the logger object
	 * @author Siri Kumar Puttagunta
	 * @param Logger object
	 */
	public static void setLoggerObject(Logger _logObject) {
		_log = _logObject;
	}

	/**
	 * Function to return the Logger object
	 * @author Siri Kumar Puttagunta
	 * @return Logger object*/
	public static Logger getLoggerObject() {
		return _log;
	}

	/**
	 * Function to print logs in the execution logs file
	 * @author Siri Kumar Puttagunta
	 * @param Globals.Log object, String message to be printed*/
	public static void log(Globals.Log level, String message) {
		if (!isLoggingInitialized) {
			setLoggerConfiguration();
		}
		if (_log != null) {
			switch (Globals.Log.valueOf(level.toString())) {
			case DEBUG:
				_log.debug(message);
				break;
			case ERROR:
				_log.error(message);
				break;
			case FATAL:
				_log.fatal(message);
				break;
			case WARN:
				_log.warn(message);
				break;
			case INFO:
				_log.info(message);
				break;
			default:
				_log.info(message);
				break;
			}
		} else {
			System.err
					.println("Call SeleLogger.setLoggerObject(_log) method in your class to create logs");
		}
	}

	/**
	 * Function to print exception details in the execution logs file
	 * @author Siri Kumar Puttagunta
	 * @param Globals.Log - object, Exception's object*/
	public static void log(Globals.Log level, Throwable exception) {
		if (!isLoggingInitialized) {
			setLoggerConfiguration();
		}
		if (_log != null) {
			switch (Globals.Log.valueOf(level.toString())) {
			case DEBUG:
				_log.debug(exception.getMessage(), exception);
				break;
			case ERROR:
				_log.error(exception.getMessage(), exception);
				break;
			case FATAL:
				_log.fatal(exception.getMessage(), exception);
				break;
			case WARN:
				_log.warn(exception.getMessage(), exception);
				break;
			case INFO:
				_log.info(exception.getMessage(), exception);
				break;
			default:
				_log.info(exception.getMessage(), exception);
				break;
			}
		} else {
			System.err
					.println("Call AELogger.setLoggerObject(_log) method in your class to create logs");
		}
	}
    
	/**
	 * Function to print exception details in the execution logs file
	 * @author Siri Kumar Puttagunta
	 * @param Globals.Log - object, Exception's object*/
	public static void log(Object obj) {
		if (!isLoggingInitialized) {
			setLoggerConfiguration();
		}
		if (_log != null) {
			Appender _appender = new AppenderSkeleton() {

				@Override
				public boolean requiresLayout() {
					// TODO Auto-generated method stub
					return false;
				}

				@Override
				public void close() {
					// TODO Auto-generated method stub

				}

				@Override
				protected void append(LoggingEvent arg0) {
					// TODO Auto-generated method stub

				}
			};
			_log.addAppender(_appender);
		} else {
			System.err
					.println("Call SeleLogger.setLoggerObject(_log) method in your class to create logs");
		}
	}

}
