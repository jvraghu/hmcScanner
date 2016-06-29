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

import java.awt.BasicStroke;
import java.awt.Color;
import java.awt.Font;
import java.awt.FontMetrics;
import java.awt.Graphics2D;

import javax.swing.JPanel;


public class PoolGraph extends JPanel {
	
	/**
	 * 
	 */
	private static final long serialVersionUID = 4412281173247423742L;
	
	private static int UPPER_GRAPH_SPACE	= 40;
	private static int LOWER_GRAPH_SPACE	= 120;
	private static int RIGHT_GRAPH_SPACE	= 20;
	private static int LEFT_GRAPH_SPACE		= 35;
	
	private String title = null;
	private float maxValue = 0;
	private String label[] = null;
	private int step = 0;
	
	private static byte		PC 		= 0;
	private static byte		SIZE 	= 1;
	private static byte		AVAIL 	= 2;
	private static byte		NUM_VAR	= 3;
	
	private float data[][]=null;
	
	private static byte		MIN		= 0;
	private static byte		AVG		= 1;
	private static byte		MAX		= 2;
	private static byte		DEVSTD	= 3;
	private static byte		P90		= 4;
	private static byte		NUM		= 5;
	private static byte		M		= 6;
	private static byte		S		= 7;
	private static byte		NUM_ST	= 8;
	
	private float stats[][]=null;
	
	private Color color[] = null;
	private String name[] = null;
	
	
	private static final float dash[] = {8,8};
	private static BasicStroke lineStroke = new BasicStroke(2.0f,BasicStroke.CAP_ROUND,BasicStroke.JOIN_ROUND);
	private static BasicStroke dotStroke = new BasicStroke(4f,BasicStroke.CAP_ROUND,BasicStroke.JOIN_ROUND,1f,dash,2);

	
	
	private void computeStats() {
		int i,j;
		
		stats = new float[NUM_VAR][NUM_ST];
		
		for (i=0; i<NUM_VAR; i++) {
			stats[i][MIN]=Float.MAX_VALUE;
			stats[i][MAX]= -1;
			stats[i][AVG]=0;
			stats[i][NUM]=0;
			
			for (j=0; j<data[i].length; j++) {
				if (data[i][j]<0)
					continue;
				
				if (stats[i][MIN]>data[i][j]) stats[i][MIN]=data[i][j];
				if (stats[i][MAX]<data[i][j]) stats[i][MAX]=data[i][j];
				stats[i][AVG]+=data[i][j];
				
				// Standard deviation & counter
				if (stats[i][NUM] == 0) {
					stats[i][M] = data[i][j];
					stats[i][S] = 0;
					stats[i][NUM] = 1;
				} else {
					float prev_m = stats[i][M];
					float prev_s = stats[i][S];
					stats[i][NUM]++;
					stats[i][M] = prev_m + ( data[i][j] - prev_m ) / stats[i][NUM];
					stats[i][S] = prev_s + ( data[i][j] - prev_m ) * ( data[i][j] - stats[i][M] );
				}
				
				
			}
			
			stats[i][AVG] = stats[i][AVG] / stats[i][NUM];
			stats[i][DEVSTD] = (float)Math.sqrt( stats[i][S] / (stats[i][NUM]-1));
			stats[i][P90] = (float)(1.282 * stats[i][DEVSTD] + stats[i][AVG]);
			
			// Format data with only two digits
			stats[i][MIN] = 1f*(int)(stats[i][MIN]*100)/100;
			stats[i][MAX] = 1f*(int)(stats[i][MAX]*100)/100;
			stats[i][AVG] = 1f*(int)(stats[i][AVG]*100)/100;
			stats[i][DEVSTD] = 1f*(int)(stats[i][DEVSTD]*100)/100;
			stats[i][P90] = 1f*(int)(stats[i][P90]*100)/100;
		}
	}
	
	
	public PoolGraph(String title, float pc[], float size[], float avail[], String label[], int step) {
		super();
		setup(title,pc,size,avail,label,step);
	}
	

	public PoolGraph(String title, float pc[], float size[], String label[], int step) {
		super();
		setup(title,pc,size,size,label,step);
	}
	
	
	private void setup(String title, float pc[], float size[], float avail[], String label[], int step) {
		this.title = title;
		this.label=label;
		this.step=step;
		
		data = new float[NUM_VAR][];
		data[PC]=pc;
		data[SIZE]=size;
		data[AVAIL]=avail;

		
		computeStats();
		
		for (int i=0; i<NUM_VAR; i++)
			if (stats[i][MAX]>maxValue)
				maxValue = stats[i][MAX];
		
		color = new Color[NUM_VAR];
		color[SIZE] = Color.GREEN;
		color[PC] = Color.BLUE;
		color[AVAIL] = Color.YELLOW;
		
		name = new String[NUM_VAR];
		name[SIZE] 	= "Pool size";
		name[PC] 	= "Processor consumed";
		name[AVAIL] = "Available CPU in parent pool";
	}
	
	
	private void drawGraph(java.awt.Graphics g, float f[], Color color, int xsize, int ysize) {
		int goodx=-1;				// good value xy 
		int goody=-1;				// good value xy 
		int i;
		int x,y;
		
		Graphics2D  g2 = (Graphics2D)g;
		
		goodx=-1;				
		goody=-1;
		g2.setColor(color);
		g2.setStroke(lineStroke);
		for (i=0; i<f.length; i++) {
			
			// Skip invalid data
			if (f[i]<0) {
				goodx=-1;				
				goody=-1;
				continue;
			}
			
			x = LEFT_GRAPH_SPACE + (int)(1f * i * (xsize-LEFT_GRAPH_SPACE-RIGHT_GRAPH_SPACE) / f.length);
			y = ysize-LOWER_GRAPH_SPACE - (int)(1f * f[i] * (ysize-UPPER_GRAPH_SPACE-LOWER_GRAPH_SPACE) / maxValue);
			
			if (goodx>=0)
				g2.drawLine(goodx,goody,x,y);
			
			goodx=x;
			goody=y;		
		}
		
	}
	
	
	private void fillOverGraph(java.awt.Graphics g, float f[], Color color, int xsize, int ysize) {
		int goodx=-1;				// good value xy 
		int goody=-1;				// good value xy 
		int i;
		int x,y;
		int px[]=new int[4];
		int py[]=new int[4];
		
		Graphics2D  g2 = (Graphics2D)g;
		
		goodx=-1;				
		goody=-1;
		g2.setColor(color);
		g2.setStroke(lineStroke);
		for (i=0; i<f.length; i++) {
			
			// Skip invalid data
			if (f[i]<0) {
				goodx=-1;				
				goody=-1;
				continue;
			}
			
			x = LEFT_GRAPH_SPACE + (int)(1f * i * (xsize-LEFT_GRAPH_SPACE-RIGHT_GRAPH_SPACE) / f.length);
			y = ysize-LOWER_GRAPH_SPACE - (int)(1f * f[i] * (ysize-UPPER_GRAPH_SPACE-LOWER_GRAPH_SPACE) / maxValue);
			
			if (goodx>=0) {
				px[0]=goodx; py[0]=UPPER_GRAPH_SPACE;
				px[1]=goodx; py[1]=goody;
				px[2]=x; py[2]=y;
				px[3]=x; py[3]=UPPER_GRAPH_SPACE;
				g.drawPolygon(px,py,4);
				g.fillPolygon(px,py,4);;
			}
							
			goodx=x;
			goody=y;		
		}
		
	}
	
	
	private void fillUnderGraph(java.awt.Graphics g, float f[], Color color, int xsize, int ysize) {
		int goodx=-1;				// good value xy 
		int goody=-1;				// good value xy 
		int i;
		int x,y;
		int px[]=new int[4];
		int py[]=new int[4];
		
		Graphics2D  g2 = (Graphics2D)g;
		
		goodx=-1;				
		goody=-1;
		g2.setColor(color);
		g2.setStroke(lineStroke);
		for (i=0; i<f.length; i++) {
			
			// Skip invalid data
			if (f[i]<0) {
				goodx=-1;				
				goody=-1;
				continue;
			}				
			
			x = LEFT_GRAPH_SPACE + (int)(1f * i * (xsize-LEFT_GRAPH_SPACE-RIGHT_GRAPH_SPACE) / f.length);
			y = ysize-LOWER_GRAPH_SPACE - (int)(1f * f[i] * (ysize-UPPER_GRAPH_SPACE-LOWER_GRAPH_SPACE) / maxValue);
			
			if (goodx>=0) {
				px[0]=goodx; py[0]=ysize-LOWER_GRAPH_SPACE;
				px[1]=goodx; py[1]=goody;
				px[2]=x; py[2]=y;
				px[3]=x; py[3]=ysize-LOWER_GRAPH_SPACE;
				g.drawPolygon(px,py,4);
				g.fillPolygon(px,py,4);
			}
							
			goodx=x;
			goody=y;		
		}
		
	}
	
	
	private void axes(java.awt.Graphics g, int xsize, int ysize) {
		Graphics2D  g2 = (Graphics2D)g;
		float pattern[]=new float[2];
		pattern[0]=10f;
		pattern[1]=5f;
		int x, y;
		int i;
		float f;
		FontMetrics metrics;
		int size;
		Font font = new Font("SansSerif", Font.BOLD, 10);
		String s;
		
		g2.setColor(Color.BLACK);
		g2.setStroke(lineStroke);		
		g2.drawLine(LEFT_GRAPH_SPACE,UPPER_GRAPH_SPACE,LEFT_GRAPH_SPACE,ysize-LOWER_GRAPH_SPACE+2);
		g2.drawLine(LEFT_GRAPH_SPACE-2,ysize-LOWER_GRAPH_SPACE,xsize-RIGHT_GRAPH_SPACE+2,ysize-LOWER_GRAPH_SPACE);
		
		g2.setPaint(Color.BLACK); 
		g2.setFont(font);
		metrics = getFontMetrics( font );
		g2.setStroke(new BasicStroke(1.0f,BasicStroke.CAP_ROUND,BasicStroke.JOIN_ROUND,5.0f,pattern,0));
		for (i=1; i<10; i++) {
			y = ysize-LOWER_GRAPH_SPACE - (int)(1f * i*(maxValue/10) * (ysize-UPPER_GRAPH_SPACE-LOWER_GRAPH_SPACE) / maxValue);
			g2.drawLine(LEFT_GRAPH_SPACE,y,xsize-RIGHT_GRAPH_SPACE,y);
			f = (int)(1f * i*(maxValue/10) *100)/100f;
			s = Float.toString(f);
			size = metrics.stringWidth(s);
			g2.drawString(s, LEFT_GRAPH_SPACE-4-size, y);
		}
		
		
		g2.setPaint(Color.BLACK); 
		g2.setFont(font);
		g2.setStroke(new BasicStroke(1.0f,BasicStroke.CAP_ROUND,BasicStroke.JOIN_ROUND,5.0f,pattern,0));
		metrics = getFontMetrics( font );
		f = step;
		i=0;
		while (f<data[PC].length) {
			x = LEFT_GRAPH_SPACE + (int)(1f * f * (xsize-LEFT_GRAPH_SPACE-RIGHT_GRAPH_SPACE) / data[PC].length);
			g2.drawLine(x,UPPER_GRAPH_SPACE,x,ysize-LOWER_GRAPH_SPACE);
			size = metrics.stringWidth(label[i]);
			g2.drawString(label[i], x-size/2, ysize-LOWER_GRAPH_SPACE+4+font.getSize());
			f += step;
			i++;
		}
		
		

		
	}
	
	
	private void title(java.awt.Graphics g, int xsize, int ysize) {
		FontMetrics metrics;
		Font font = new Font("SansSerif", Font.BOLD, 16);
		int size;
		Graphics2D  g2 = (Graphics2D)g;
		
		metrics = getFontMetrics( font );
		size = metrics.stringWidth(title);
		
		g2.setFont(font);
		g2.setColor(Color.BLACK);
		g2.drawString(title, xsize/2-size/2, font.getSize()+2);
	}
	
	
	private void summary(java.awt.Graphics g, int xsize, int ysize) {
		FontMetrics metrics;
		Font font = new Font("SansSerif", Font.PLAIN, 12);
		Graphics2D  g2 = (Graphics2D)g;
		int y;
		int px[]=new int[4];
		int py[]=new int[4];
		int i,j;
		int largerLabel=0;
		int largerData[] = new int[NUM_ST];
		int labelTab = 0;
		int statsTab[] = new int[NUM_ST];
		String statsLabel[] = { "Min   ", "Avg   ", "Max   ", "StdDev", "90° Perc" };
		int size;
 		
		g2.setStroke(lineStroke);
		g2.setFont(font);
		g2.setColor(Color.BLACK);
		metrics = getFontMetrics( font );
		
		for (i=0; i<NUM_VAR; i++) {
			if (name[i].length()>name[largerLabel].length())
				largerLabel=i;
			
			for (j=0; j<NUM_ST; j++)
				if (Float.toString(stats[i][j]).length()>Float.toString(stats[largerData[j]][j]).length())
					largerData[j] = i;
		}
		
		
		labelTab = 30;
		statsTab[0] = labelTab + metrics.stringWidth(name[largerLabel]) + 20;
		for (i=1; i<=P90; i++) {
			size = metrics.stringWidth(Float.toString(stats[largerData[i]][i]));
			if (metrics.stringWidth(statsLabel[i]) > size)
				size = metrics.stringWidth(statsLabel[i]);
			statsTab[i] = statsTab[i-1] + size + 10;
		}
		
		
		
		y = ysize-LOWER_GRAPH_SPACE+font.getSize()+30;
		
		g2.setColor(Color.BLACK);
		for (i=0; i<=P90; i++) 
			g2.drawString(statsLabel[i], statsTab[i], y);
		y += font.getSize()+4;
		
		for (i=0; i<NUM_VAR; i++) {
			
			px[0] = 10; py[0] = y;
			px[1] = 20; py[1] = y;
			px[2] = 20; py[2] = y-10;
			px[3] = 10; py[3] = y-10;
			
			g2.setColor(Color.BLACK);
			g.drawPolygon(px,py,4);
			g2.setColor(color[i]);
			g.fillPolygon(px,py,4);
			g2.setColor(Color.BLACK);
			g2.drawString(name[i], labelTab, y);
			for (j=0; j<=P90; j++)
				g2.drawString(Float.toString(stats[i][j]),statsTab[j], y);
			y += font.getSize()+4;
		}
		
	}
	
	
	private void drawP90(java.awt.Graphics g, float p90, Color color, int xsize, int ysize) {
		int y;
		
		Graphics2D  g2 = (Graphics2D)g;
		

		g2.setColor(color);
		g2.setStroke(dotStroke);
		y = ysize-LOWER_GRAPH_SPACE - (int)(1f * p90 * (ysize-UPPER_GRAPH_SPACE-LOWER_GRAPH_SPACE) / maxValue);
		g2.drawLine(LEFT_GRAPH_SPACE,y,xsize-RIGHT_GRAPH_SPACE,y);		
	}

	
	public void paint(java.awt.Graphics g) {	
		
		// Start painting panel
		super.paintComponent(g);
		
		int ysize = getHeight();		// vertical size
		int xsize = getWidth();			// horizontal size
		
		g.setColor(java.awt.Color.lightGray);
		g.fillRect(0,0,xsize,ysize);	
		
		
		// Title
		title(g,xsize,ysize);
		
		// Available
		fillUnderGraph(g,data[AVAIL],color[AVAIL],xsize,ysize);
	
		// Size
		drawGraph(g,data[SIZE],color[SIZE],xsize,ysize);
		
		// Processor Consumed
		drawGraph(g,data[PC],color[PC],xsize,ysize);
		drawP90(g,stats[PC][P90],color[PC],xsize,ysize);
		
		// Axes
		axes(g,xsize,ysize);
		
		// Summary
		summary(g,xsize,ysize);
		

	}

}
