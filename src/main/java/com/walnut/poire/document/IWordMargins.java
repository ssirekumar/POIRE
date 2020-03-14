package com.walnut.poire.document;

public interface IWordMargins {
	
	
			
	public static enum NORMAL{
		BottomMargin(1440), //this value is converted from 1-inch to twip
		TopMargin(1440),
		LeftMargin(1440),
		RightMargin(1440);
		private final int value;
		private NORMAL(int value){
			this.value = value;	
		}
		public int getValue() {
			return this.value;
		}
	}
	
	public static enum NARROW {
		BottomMargin(720), //this value is converted from 1-inch to twip
		TopMargin(720),
		LeftMargin(720),
		RightMargin(720);
		private final int value;
		private NARROW(int value){
			this.value = value;	
		}
		public int getValue() {
			return this.value;
		}
	}
	
	public static enum MODERATE{
		BottomMargin(1440), //this value is converted from 1-inch to twip
		TopMargin(1440),
		LeftMargin(1080),
		RightMargin(1080);
		private final int value;
		private MODERATE(int value){
			this.value = value;	
		}
		public int getValue() {
			return this.value;
		}
	}
	
	public static enum WIDE{
		BottomMargin(1440), //this value is converted from 1-inch to twip
		TopMargin(1440),
		LeftMargin(2880),
		RightMargin(2880);
		private final int value;
		private WIDE(int value){
			this.value = value;	
		}
		public int getValue() {
			return this.value;
		}
	}
	
	public static enum OFFICE_2003_DEFAULT{
		BottomMargin(1440), //this value is converted from 1-inch to twip
		TopMargin(1440),
		LeftMargin(1800),
		RightMargin(1800);
		private final int value;
		private OFFICE_2003_DEFAULT(int value){
			this.value = value;	
		}
		public int getValue() {
			return this.value;
		}
	}

	

	
		
	
}
