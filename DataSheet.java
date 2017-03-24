/*
   Copyright 2016 Federico Vagnini (vagnini@it.ibm.com)

   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
*/

package hmcScanner;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintWriter;

import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.format.Orientation;
import jxl.format.VerticalAlignment;
import jxl.write.Formula;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.NumberFormats;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableImage;
import jxl.write.WritableSheet;
import jxl.write.WriteException;

public class DataSheet {
	private static final byte MIN_ROW_LENGTH 	= 10; // a row is allocated as multiple of this value
	private static final byte MIN_ROWS 			= 10; // every time allocate this set of rows
	

	// Cell formats
	
	private static int		NONE			= 0;
	private static int		BOLD			= 1;
	private static int		CENTRE			= 1 << 1;
	private static int		RIGHT			= 1 << 2;
	private static int		LEFT			= 1 << 3;
	private static int		VCENTRE			= 1 << 4;
	private static int		B_TOP_MED		= 1 << 5;
	private static int		B_BOTTOM_MED	= 1 << 6;
	private static int		B_LEFT_MED		= 1 << 7;
	private static int		B_RIGHT_MED		= 1 << 8;
	private static int		B_ALL_MED		= 1 << 9;
	private static int		B_TOP_LOW		= 1 << 10;
	private static int		B_BOTTOM_LOW	= 1 << 11;
	private static int		B_LEFT_LOW		= 1 << 12;
	private static int		B_RIGHT_LOW		= 1 << 13;
	private static int		B_ALL_LOW		= 1 << 14;
	private static int		GRAY_25			= 1 << 15;
	private static int		GREEN			= 1 << 16;
	private static int		WRAP			= 1 << 17;
	private static int		DIAG45			= 1 << 18;
	private static int		BLACK			= 1 << 19;
	private static int		YELLOW			= 1 << 20;
	private static int		RED				= 1 << 21;
	
	private int nrows = 0;					// valid rows
	private int ncolumns[] = null;			// for each row, the number of valid columns
	private int colSize[] = null;			// size of each column
	
	private Cell sheet[][] = null;			// the sheet we are building
	
	private char			separator = ',';
	
	
	
	// CELL TYPES
	private static final byte		LABEL	= 0;
	private static final byte		FLOAT	= 1;
	private static final byte		INTEGER	= 2;
	private static final byte		IMAGE	= 3;
	private static final byte		FORMULA	= 4;
	private static final byte		VOID	= 5;
	
	
	public class Cell {
		// This class represents a single cell with its own generic descriptors
		
		private byte		type = LABEL;
		private String		string = null;
		private double		number = 0;
		private int			map = 0;
		private int			width = 0;
		private int			height = 0;
		private int			merge_to_row = -1;		// no merge
		private int			merge_to_col = -1;		// no merge
		
		
		public void setType(byte v) { type = v; }
		public byte getType() { return type; }
		public void setString(String s) { string = s; };
		public String getString() { return string; };
		public void setNumber(double d) { number = d; };
		public double getNumber() { return number; };
		public void setMap(int m) { map = m; };
		public int getMap() { return map; };
		public void setWidth(int w) { width = w; };
		public int getWidth() { return width; };
		public void setHeight(int w) { height = w; };
		public int getHeight() { return height; }
		public int getMerge_to_row() { return merge_to_row;	}
		public void setMerge_to_row(int merge_to_row) {this.merge_to_row = merge_to_row;}
		public int getMerge_to_col() { return merge_to_col;	}
		public void setMerge_to_col(int merge_to_col) {	this.merge_to_col = merge_to_col; };
	}
	
	
	public DataSheet() {
		sheet = null;
		nrows = 0;
		ncolumns = null;
		colSize = null;
	}
	
	public static void main(String[] args) {
		DataSheet sd = new DataSheet();
		/*
		sd.getCell(1,2);
		sd.getCell(0,0);
		sd.getCell(0, 20);
		sd.getCell(20, 0);
		sd.getCell(100, 100);
		*/
		
	}
	
	
	public void createExcelSheet(WritableSheet excelSheet) {
		int row, col;
		
		// Create cells
		for (row=0; row<nrows; row++)
			for (col=0; col<ncolumns[row]; col++) {
				if (sheet[row][col]==null)
					continue;
						
				try {
					if (sheet[row][col].getMerge_to_col()>=0 && sheet[row][col].getMerge_to_row()>=0)
						excelSheet.mergeCells(col, row, sheet[row][col].getMerge_to_col(), sheet[row][col].getMerge_to_row());
					
					switch (sheet[row][col].getType()) {
						case LABEL:
							excelSheet.addCell(new Label(col,row,sheet[row][col].getString(),getExcelFormat(row, col)));
							break;
						case FLOAT:
						case INTEGER:
							excelSheet.addCell(new Number(col,row,sheet[row][col].getNumber(),getExcelFormat(row, col)));
							break;
						case IMAGE:
							File f = new File(sheet[row][col].getString());
							WritableImage imgobj=new WritableImage(col, row, sheet[row][col].getWidth(), sheet[row][col].getHeight(), f);
							excelSheet.addImage(imgobj);
							break;
						case VOID:
							break;
						case FORMULA:
							excelSheet.addCell( new Formula(col, row, sheet[row][col].getString(),getExcelFormat(row, col)) );
							break;
						default:
							break;
					}
				} catch (WriteException we) {};
			}
		
		// Resize columns
		for (col=0; colSize!=null && col<colSize.length; col++)
			if (colSize[col]>0)
				excelSheet.setColumnView(col, colSize[col]);
	}
	
	
	public void createCSVSheet(String fileName) {
		int row, col;
		PrintWriter csv=null;
		String s;
		double d;
		int i;
		
		try {
			csv = new PrintWriter(
					new FileOutputStream(
					new File(fileName)));
		}
		catch (IOException e) { return; }
		
		// Create cells
		for (row=0; row<nrows; row++) {
			for (col=0; col<ncolumns[row]; col++) {
				if (sheet[row][col]!=null) {

					switch (sheet[row][col].getType()) {
						case LABEL:
							s = sheet[row][col].getString();
							if (s!=null)
								csv.print(s);
							break;
						case INTEGER:
							i = (int)(sheet[row][col].getNumber());
							csv.print(i);
							break;
						case FLOAT:
							d = sheet[row][col].getNumber();
							// Only keep 2 digits: avoid rounding errors
							d = 1d*(int)(d*100)/100;
							csv.print(d);
							break;
						case IMAGE:
							break;
						case VOID:
							break;
						case FORMULA:
							break;
						default:
							break;
					}
					
				}
				
				if (col<ncolumns[row]-1)
					csv.print(separator);
			}
			csv.println();
		}
		
		csv.close();
	}
	
	
	public void setSeparator(char separator) {
		this.separator = separator;
	}

	public void createHTMLSheet(String fileName) {
		PrintWriter html=null;
		int row, col;
		String s;
		int map;
		int merge;
		double d;

		
		try {
			html = new PrintWriter(
					new FileOutputStream(
					new File(fileName)));
		}
		catch (IOException e) { return; }
		
		html.println("<HTML>\n" +
				"<HEAD>\n" +
				"<style>\n" +
				"	table{border-collapse:collapse}\n" +
				"	th,td" +
				"{border:1px solid black;border-collapse:collapse;padding:7px;font-family:calibri;font-size:12px;}\n" +
				"</style>\n</HEAD>\n" +
				"<BODY style=\"background-color:#e6e6e6;\">\n" +
				"<TABLE>\n");
		
		
		// Create cells
		for (row=0; row<nrows; row++) {
			if (ncolumns[row]==0) {
				html.println("<TR style=\"border:0px\">");
				html.println("\t<TD style=\"border:0px\"></TD>");
				html.println("</TR>");
			} else {			
				html.println("<TR>");
				for (col=0; col<ncolumns[row]; col++) {
					
					if (sheet[row][col]==null) {
						html.print("\t<TD style=\"border:0px\">");				
						html.print("</TD>");
					} else if (sheet[row][col].getType()==VOID) {
						;
					} else {					
						map = sheet[row][col].getMap();
						s="";
						if ( (map&GREEN)!=0 ) 	s+="background-color:LightGreen;";
						if ( (map&BLACK)!=0 ) 	s+="background-color:Black;";
						if ( (map&YELLOW)!=0 ) 	s+="background-color:Yellow;";
						if ( (map&RED)!=0 ) 	s+="background-color:Red;";
						
						if ( (map&BOLD) !=0 )	s+="font-weight:bold;";
						if ( (map&CENTRE) !=0 )	s+="text-align:center;";
						if ( (map&LEFT) !=0 )	s+="text-align:left;";
						if ( (map&RIGHT) !=0 )	s+="text-align:right;";
						
						if (s.length()>0)
							s = " style=\""+s+"\"";
						
						merge = sheet[row][col].getMerge_to_col();
						if (merge>0 && merge>col)
							s+=" colspan=\""+(merge-col+1)+"\"";
						merge = sheet[row][col].getMerge_to_row();
						if (merge>0 && merge>row)
							s+=" rowspan=\""+(merge-row+1)+"\"";
						
						html.print("\t<TD"+s+">");
						
						/*
						if (s.length()>0)
							html.print("\t<TD style=\""+s+"\">");
						else
							html.print("\t<TD>");
						*/
						
						switch (sheet[row][col].getType()) {
							case LABEL:
								s = sheet[row][col].getString();
								if (s!=null)
									html.print(s);
								break;
							case INTEGER:
								html.print((int)(sheet[row][col].getNumber()));
								break;
							case FLOAT:
								d = sheet[row][col].getNumber();
								// Only keep 2 digits: avoid rounding errors
								d = 1d*(int)(d*100)/100;
								html.print(d);
								break;
							case IMAGE:
								InputStream input = null;
								OutputStream output = null;
								try {
									input = new FileInputStream(sheet[row][col].getString());
									String dest = new File(fileName).getParent() + File.separatorChar + new File(sheet[row][col].getString()).getName();
									//String dest = new File(fileName).getParent() + File.pathSeparatorChar + new File(fileName).getName();
									output = new FileOutputStream(dest);
									byte[] buf = new byte[1024];
									int bytesRead;
									while ((bytesRead = input.read(buf)) > 0) {
										output.write(buf, 0, bytesRead);
									}	
									input.close();
									output.close();
									html.print("<IMG SRC=\""+new File(sheet[row][col].getString()).getName()+"\">");
								} 
								catch (IOException e) {}
								break;
							case FORMULA:
								// Nothing for now....
								break;
							default:
								break;
						}
						
						html.print("</TD>");
					} 
				}
				html.println("\n</TR>");
			}
		}
		
		html.println("</TABLE>\n</BODY>\n</HTML>");
		
		html.close();
		
	}
	
	
	private WritableCellFormat getExcelFormat(int row, int col) {
		Cell c = sheet[row][col];
		int map = c.getMap();
		byte type = c.getType();
		WritableCellFormat wcf;
		WritableFont wf;
		
		if ( (map & BOLD) != 0 )
			wf = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD, true);
		else
			wf = new WritableFont(WritableFont.ARIAL, 10);
		
		switch (type) {
			case LABEL:
				wcf = new WritableCellFormat(wf); break;
			case FLOAT:
			case FORMULA:
				wcf = new WritableCellFormat(wf,NumberFormats.FORMAT3);	break;
			case INTEGER:
				wcf = new WritableCellFormat(wf,NumberFormats.INTEGER);	break;
			default:
				wcf = null; break;
		}
		
		try {
			if ( (map&CENTRE)!=0 ) 		wcf.setAlignment(Alignment.CENTRE);
			if ( (map&RIGHT)!=0 ) 		wcf.setAlignment(Alignment.RIGHT);
			if ( (map&LEFT)!=0 ) 		wcf.setAlignment(Alignment.LEFT);
			if ( (map&VCENTRE)!=0 ) 	wcf.setVerticalAlignment(VerticalAlignment.CENTRE);
			
			if ( (map&B_TOP_MED)!=0 ) 	wcf.setBorder(Border.TOP, BorderLineStyle.MEDIUM);
			if ( (map&B_BOTTOM_MED)!=0 ) wcf.setBorder(Border.BOTTOM, BorderLineStyle.MEDIUM);
			if ( (map&B_LEFT_MED)!=0 ) 	wcf.setBorder(Border.LEFT, BorderLineStyle.MEDIUM);
			if ( (map&B_RIGHT_MED)!=0 ) wcf.setBorder(Border.RIGHT, BorderLineStyle.MEDIUM);
			
			if ( (map&B_ALL_MED)!=0 ) {
				wcf.setBorder(Border.RIGHT, BorderLineStyle.MEDIUM);
				wcf.setBorder(Border.LEFT, BorderLineStyle.MEDIUM);
				wcf.setBorder(Border.TOP, BorderLineStyle.MEDIUM);
				wcf.setBorder(Border.BOTTOM, BorderLineStyle.MEDIUM);
			}
			
			if ( (map&B_TOP_LOW)!=0 ) 	wcf.setBorder(Border.TOP, BorderLineStyle.THIN);
			if ( (map&B_BOTTOM_LOW)!=0 ) wcf.setBorder(Border.BOTTOM, BorderLineStyle.THIN);
			if ( (map&B_LEFT_LOW)!=0 ) 	wcf.setBorder(Border.LEFT, BorderLineStyle.THIN);
			if ( (map&B_RIGHT_LOW)!=0 ) wcf.setBorder(Border.RIGHT, BorderLineStyle.THIN);
			
			if ( (map&B_ALL_LOW)!=0 ) {
				wcf.setBorder(Border.RIGHT, BorderLineStyle.THIN);
				wcf.setBorder(Border.LEFT, BorderLineStyle.THIN);
				wcf.setBorder(Border.TOP, BorderLineStyle.THIN);
				wcf.setBorder(Border.BOTTOM, BorderLineStyle.THIN);
			}
			
			if ( (map&GRAY_25)!=0 ) wcf.setBackground(Colour.GRAY_25);
			if ( (map&GREEN)!=0 ) wcf.setBackground(Colour.LIGHT_GREEN);
			if ( (map&BLACK)!=0 ) wcf.setBackground(Colour.BLACK);
			if ( (map&YELLOW)!=0 ) wcf.setBackground(Colour.YELLOW);
			if ( (map&RED)!=0 ) wcf.setBackground(Colour.RED);
			
			if ( (map&WRAP)!=0 ) wcf.setWrap(true);
			
			if ( (map&DIAG45)!=0 ) wcf.setOrientation(Orientation.PLUS_45);
		} catch (WriteException we) {
			System.out.println("DataSheet.getExcelFormat: WriteException. Keep going...");
		}
		
		return wcf;
	}
	
	
	public void setColSize(int col, int size) {
		if (colSize==null) {
			colSize = new int[sheet[0].length];
		}
		if (col>=colSize.length)
			return;
		colSize[col]=size;
	}
	
	
	public void addFormula(int col, int row, String formula, int map) {
		Cell c = getCell(row, col);
		c.setString(formula);
		c.setMap(map);
		c.setType(FORMULA);
	
		return;
	}
	
	
	
	public int addLabel(int col, int row, String s, int map) {
		Cell c = getCell(row, col);
		c.setString(s);
		c.setMap(map);
		c.setType(LABEL);
		
		if (s==null)
			return 0;
		return s.length();
	}
	
	public int addLabel(int col, int row, String s[], int index, int map) {
		Cell c = getCell(row, col);
		c.setMap(map);
		c.setType(LABEL);
		if (s==null || index>=s.length) {
			c.setString(null);
			return 0;
		}
		
		c.setString(s[index]);
		if (s[index]==null)
			return 0;
		return s[index].length();
	}
	
	public void addMultipleLabelsWrap(int col, int row, String s[], int map) {
		Cell c = getCell(row, col);
		c.setMap(map|WRAP);
		c.setType(LABEL);
		if (s==null || s.length==0){
			c.setString(null);
			return;
		}
		
		if (s.length==1) {
			c.setString(s[0]);
			return;
		}
			
		String result = s[0];
		
		for (int i=1; i<s.length; i++)
			result = result + "\n" + s[i];
		c.setString(result);	
	}
	
	
	public void addFloat(int col, int row, String s[], int index, int map) {
		addNumber(col, row, s, index, map, FLOAT);
	}
	public void addInteger(int col, int row, String s[], int index, int map) {
		addNumber(col, row, s, index, map, INTEGER);
	}
	
	public void addFloat(int col, int row, double d, int map) {
		addNumber(col, row, d, map, FLOAT);
	}
	public void addInteger(int col, int row, double d, int map) {
		addNumber(col, row, d, map, INTEGER);
	}
	
	public void addFloatDiv1024(int col, int row, String s[], int index, int map) {
		addNumberDiv1024(col, row, s, index, map, FLOAT);
	}
	public void addIntegerDiv1024(int col, int row, String s[], int index, int map) {
		addNumberDiv1024(col, row, s, index, map, INTEGER);
	}
	
	
	public void addPicture(int col, int row, int width, int height, String file ) {
		File f = new File(file);
		if (!f.exists()) 
			return;
		
		Cell c = getCell(row, col);
		c.setString(file);
		c.setType(IMAGE);	
		c.setWidth(width);
		c.setHeight(height);
	}
	
	
	private void addNumber(int col, int row, String s[], int index, int map, byte type)  {	
		if (s==null || index>=s.length || s[index]==null)
			return;
		
		double d; 
		
		try {
			d = Double.parseDouble(s[index]);
			
			// Only keep 2 digits: avoid rounding errors
			d = 1d*(int)(d*100)/100;
		} catch (NumberFormatException nfe) {
			Cell c = getCell(row, col);
			c.setMap(map);
			c.setType(LABEL);
			c.setString("NaN");
			return;
		}	
				
		Cell c = getCell(row, col);
		c.setMap(map);
		c.setNumber(d);
		c.setType(type);
	}
	
	
	private void addNumber(int col, int row, double d, int map, byte type) {
		Cell c = getCell(row, col);
		c.setMap(map);
		c.setNumber(d);
		c.setType(type);
	}
	
	private void addNumberDiv1024(int col, int row, String s[], int index, int map, byte type) {		
		if (s==null || index>=s.length || s[index]==null)
			return;
		
		double d; 
		
		try {
			d = Double.parseDouble(s[index]);
			d = d /1024;
			
			// Only keep 2 digits: avoid rounding errors
			d = 1d*(int)(d*100)/100;
		} catch (NumberFormatException nfe) {
			Cell c = getCell(row, col);
			c.setMap(map);
			c.setType(LABEL);
			c.setString("NaN");
			return;
		}	
				
		Cell c = getCell(row, col);
		c.setMap(map);
		c.setNumber(d);
		c.setType(type);
	}
	
	
	/*
	 * If the cell exists, return it. Otherwise create a new one in the provided position.
	 */
	private Cell getCell(int row, int col) {
		Cell c;
		Cell new_sheet[][] = null;
		int i,j;
		int sheet_width;
		
		if (row<nrows && col<ncolumns[row]) {
			// The cell is inside the currently allocated zone
			if (sheet[row][col]!=null)
				return sheet[row][col];
			c = new Cell();
			sheet[row][col] = c;
			return c;
		}
		
		if (row<nrows) {
			// The row exists but the col is bigger than current row size
			
			if (col<sheet[row].length) {
				// Sheet is big enough, just update row size
				ncolumns[row]=col+1;
				c = new Cell();
				sheet[row][col] = c;
				return c;
			}
			
			// sheet row needs to be enlarged
			if (sheet[0].length > (1+col/MIN_ROW_LENGTH)*MIN_ROW_LENGTH)
				sheet_width = sheet[0].length;
			else
				sheet_width = (1+col/MIN_ROW_LENGTH)*MIN_ROW_LENGTH;
			new_sheet = new Cell[sheet.length][sheet_width];
			for (i=0; i<nrows; i++)
				for (j=0; j<ncolumns[i]; j++)
					new_sheet[i][j] = sheet[i][j];
			sheet = new_sheet;
			ncolumns[row]=col+1;
			c = new Cell();
			sheet[row][col] = c;
			return c;
		}
		
		// A new row needs to be added. It could be the first one!!!
		
		if (nrows==0) {
			// very first cell!
			ncolumns = new int[(1+row/MIN_ROWS)*MIN_ROWS];
			nrows = row+1;
			ncolumns[row] = col+1;
			sheet = new Cell[(1+row/MIN_ROWS)*MIN_ROWS][(1+col/MIN_ROW_LENGTH)*MIN_ROW_LENGTH];
			c = new Cell();
			sheet[row][col] = c;
			return c;
		}
		
		if (row<ncolumns.length) {
			// We have enough rows   
			
			if (col<sheet[row].length) {
				// Sheet is big enough, just update row size
				nrows = row+1;
				ncolumns[row]=col+1;
				c = new Cell();
				sheet[row][col] = c;
				return c;
			}
			
			// sheet row needs to be enlarged
			if (sheet[0].length > (1+col/MIN_ROW_LENGTH)*MIN_ROW_LENGTH)
				sheet_width = sheet[0].length;
			else
				sheet_width = (1+col/MIN_ROW_LENGTH)*MIN_ROW_LENGTH;
			new_sheet = new Cell[sheet.length][sheet_width];
			for (i=0; i<nrows; i++)
				for (j=0; j<ncolumns[i]; j++)
					new_sheet[i][j] = sheet[i][j];
			sheet = new_sheet;
			nrows = row+1;
			ncolumns[row]=col+1;
			c = new Cell();
			sheet[row][col] = c;
			return c;
		}
		
		// A new set of rows needs to be allocated
		if (sheet[0].length > (1+col/MIN_ROW_LENGTH)*MIN_ROW_LENGTH)
			sheet_width = sheet[0].length;
		else
			sheet_width = (1+col/MIN_ROW_LENGTH)*MIN_ROW_LENGTH;
		new_sheet = new Cell[(1+row/MIN_ROWS)*MIN_ROWS][sheet_width];
		for (i=0; i<nrows; i++)
			for (j=0; j<ncolumns[i]; j++)
				new_sheet[i][j] = sheet[i][j];
		sheet = new_sheet;
		int new_ncolumns[] = new int[(1+row/MIN_ROWS)*MIN_ROWS];
		for (i=0; i<nrows; i++)
			new_ncolumns[i] = ncolumns[i];
		ncolumns = new_ncolumns;
		nrows = row+1;
		ncolumns[row]=col+1;
		c = new Cell();
		sheet[row][col] = c;
		return c;
	}
	
	
	// Merge cells from (colf,rowf) to (colt,rowt)
	public void mergeCells(int colf, int rowf, int colt, int rowt) {
		Cell c = getCell(rowf, colf);
		c.setMerge_to_col(colt);
		c.setMerge_to_row(rowt);
		
		// Fill with void cells
		int i,j;
		for (i=rowf; i<=rowt; i++)
			for (j=colf; j<=colt; j++) {
				c = getCell(i, j);
				c.setType(VOID);
			}
	}
	
	

}
