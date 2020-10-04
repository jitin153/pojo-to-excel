package com.jitin.pojotoexcel.exception;

public class ExcelWriterException extends RuntimeException{

	private static final long serialVersionUID = -2659626614302534517L;
	
	public ExcelWriterException() {
		super();
	}
	
	public ExcelWriterException(String msg) {
		super(msg);
	}

	public ExcelWriterException(Throwable t) {
		super(t);
	}
	
	public ExcelWriterException(String msg, Throwable t) {
		super(msg,t);
	}
}
