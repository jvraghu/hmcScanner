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

import java.util.Calendar;
import java.util.GregorianCalendar;
import java.util.Vector;

public class NewLparStatus {
	
	private static int DAYS = 365;
	private static int HOURLY_DAYS = 60;
	
	private GregorianCalendar last;
	
	
	// Data --> variable[HOURLY|DAILY|MONTHLY][items]
	private String[][]		msName = null;
	private String[][]		poolName = null;
	private boolean[][]		cap = null;

	
	private Vector<String>	msNames = new Vector<String>();
	private Vector<String>	poolNames = new Vector<String>();
	
	private static byte	HOURLY	= 0;
	private static byte DAILY	= 1;
	private static byte MONTHLY	= 2;
	
	private boolean monthlyDone = false;
	
	
	
	
	/*
	
	// Daily data
	private String[]	msDaily, poolDaily;
	private boolean[]	capDaily;
	
	private static byte	HOURLY	= 0;
	private static byte DAILY	= 1;
	private static byte FINISHED = 2;
	private byte status = HOURLY;
	
	
	// Monthly data
	private String[]	msMonthly, poolMonthly;
	private boolean[]	capMonthly;
	private GregorianCalendar[]	monthName;
	
	private static byte NUM_MONTHS = 13;		// 13 months due to start in middle of a month!!
	
	// Hourly data
	private String[]	msHourly, poolHourly;
	private boolean[]	capHourly;
	
	*/
	
	
	/*
	 * Usage:
	 * 	1) Allocate object
	 * 	2) Add hourly data using "add"
	 * 	3) endOfHourlyData()
	 * 	4) Add daily data using "add"
	 * 	5) endOfHourlyData()
	 * 	6) getMonthNames() for labels of NUM_MONTHS months. It shows first sample of the month
	 *  7) getMonthXXX() for data related to month 0 to 11
	 *  8) getMonthNum() provides the number of data days included in the month
	 */
	
	
	
	public NewLparStatus(GregorianCalendar last) {
		// Set it to 23:59 in order to get the entire day
		this.last = (GregorianCalendar)last.clone();
		this.last.set(Calendar.HOUR_OF_DAY, 23);
		this.last.set(Calendar.MINUTE, 59);
				
		allocate();
	}
	
	
	private void createMonthlyData() {
		int		poolCount[] = new int[poolNames.size()];
		int		msCount[] = new int[msNames.size()];
		int		capCount = 0;
			
		float	num = 0;
		int 	month, day;
		GregorianCalendar curr, prev;
		int i,j;
		
		curr = (GregorianCalendar)last.clone();
		prev = (GregorianCalendar)last.clone();
		
		for (day=msName[DAILY].length-1, month=12; day>=0; day--) {
			if (curr.get(Calendar.MONTH)==prev.get(Calendar.MONTH)) {
				
				if (msName[DAILY][day]!=null) {			
					for (i=0; i<msNames.size(); i++)
						if (msName[DAILY][day].equals(msName[DAILY][i]))
							msCount[i]++;
					for (i=0; i<poolNames.size(); i++)
						if (poolName[DAILY][day].equals(poolName[DAILY][i]))
							poolCount[i]++;
					if (cap[DAILY][day])
							capCount++;
					num++;
				}
			} else {
				if (num>0) {
					j=0;
					for (i=0; i<msCount.length; i++)
						if (msCount[i]>msCount[j])
							j=i;
					msName[MONTHLY][month]=msNames.elementAt(j);
					
					j=0;
					for (i=0; i<poolCount.length; i++)
						if (poolCount[i]>poolCount[j])
							j=i;
					poolName[MONTHLY][month]=poolNames.elementAt(j);
					
					if (capCount>num/2)
						cap[MONTHLY][month]=true;
				}
				
				month--;
				for (i=0; i<msCount.length; i++)
					msCount[i]=0;
				for (i=0; i<poolCount.length; i++)
					poolCount[i]=0;
				capCount=0;
				num=0;
				
				prev = (GregorianCalendar)curr.clone();
			}
			curr.add(Calendar.DATE, -1);			
		}
		
		if (num>0) {
			j=0;
			for (i=0; i<msCount.length; i++)
				if (msCount[i]>msCount[j])
					j=i;
			msName[MONTHLY][month]=msNames.elementAt(j);
			
			j=0;
			for (i=0; i<poolCount.length; i++)
				if (poolCount[i]>poolCount[j])
					j=i;
			poolName[MONTHLY][month]=poolNames.elementAt(j);
			
			if (capCount>num/2)
				cap[MONTHLY][month]=true;
		}		
		
		monthlyDone = true;
	}
	

	private void allocate() {		
		// Allocate structure
		msName		= new String[3][];
		poolName	= new String[3][];
		cap			= new boolean[3][];
		
		// Allocate HOURLY
		msName[HOURLY] 		= new String[24*30*2];
		poolName[HOURLY] 	= new String[24*30*2];
		cap[HOURLY]			 = new boolean[24*30*2];
		
		// Allocate DAILY
		msName[DAILY] 		= new String[365];
		poolName[DAILY] 	= new String[365];
		cap[DAILY]			= new boolean[365];

		// Allocate MONTHLY
		msName[MONTHLY] 		= new String[13];
		poolName[MONTHLY] 	= new String[13];
		cap[MONTHLY]			= new boolean[13];
	}
	
	
	
	private int getSlot(GregorianCalendar gc, int type) {
		float diff;
		int slot=0;
		
		diff = last.getTimeInMillis()-gc.getTimeInMillis();
		
		if (type==HOURLY)
			slot = (HOURLY_DAYS*24-1) - (int)(diff/(60*60*1000));
		else if (type==DAILY)
			slot = (DAYS-1) - (int)(diff/(24*60*60*1000));
		
		return slot;
	}
	
	
	
	
	private String getMSString(String s) {
		int i;
		
		for (i=0; i<msNames.size(); i++) {
			if (s.equals(msNames.elementAt(i)))
				break;
		}
		
		if (i<msNames.size())
			return msNames.elementAt(i);
		
		msNames.add(s);
		
		return s;
	}
	
	
	private String getPoolString(String s) {
		int i;
		
		for (i=0; i<poolNames.size(); i++) {
			if (s.equals(poolNames.elementAt(i)))
				break;
		}
		
		if (i<poolNames.size())
			return poolNames.elementAt(i);
		
		poolNames.add(s);
		
		return s;
	}
	
	
	private void addData(GregorianCalendar gc, String ms, String pool, boolean cap, byte type) {
		if (pool.equals(""))
			pool = "DefaultPool";
		
		int slot = getSlot(gc, type);
		
		if (slot<0 || slot>=msName[type].length) {
			/*
			System.out.println("Out of bound daily data insert: "+gc.get(Calendar.YEAR)+"/"+
														(gc.get(Calendar.MONTH)+1)+"/"+
														gc.get(Calendar.DAY_OF_MONTH)+" "+
														gc.get(Calendar.HOUR_OF_DAY)+":"+
														gc.get(Calendar.MINUTE)+"/"+
														gc.get(Calendar.SECOND) );
			*/
			return;
		}
		if (msName[type][slot]!=null) {
			/*
			System.out.println("Warning! Two samples added in same day: "+gc.get(Calendar.YEAR)+"/"+
														(gc.get(Calendar.MONTH)+1)+"/"+
														gc.get(Calendar.DAY_OF_MONTH)+" "+
														gc.get(Calendar.HOUR_OF_DAY)+":"+
														gc.get(Calendar.MINUTE)+"/"+
														gc.get(Calendar.SECOND) );
			*/
			;
		}
		
		msName[type][slot] = getMSString(ms);
		poolName[type][slot] = getPoolString(pool);
		this.cap[type][slot] = cap;
	}
	
	public void addHourData(GregorianCalendar gc, String ms, String pool, boolean cap) {
		addData(gc, ms, pool, cap, HOURLY);
	}
	public void addDayData(GregorianCalendar gc, String ms, String pool, boolean cap) {
		addData(gc, ms, pool, cap, DAILY);
	}
	

	
	public String[] getHourlyLabels() {
		String result[] = new String[msName[HOURLY].length];
		GregorianCalendar gc;
		
		for (int i=0; i<msName[HOURLY].length; i++) {
			gc=(GregorianCalendar)last.clone();
			gc.add(Calendar.HOUR, -i);
			result[msName[HOURLY].length-1-i] = gc.get(Calendar.YEAR)+"/"+
					(gc.get(Calendar.MONTH)+1)+"/"+
					gc.get(Calendar.DAY_OF_MONTH)+" "+
					gc.get(Calendar.HOUR_OF_DAY);
		}
		return result;
	}
	
	public String[] getDailyLabels() {
		String result[] = new String[msName[DAILY].length];
		GregorianCalendar gc;
		
		for (int i=0; i<msName[DAILY].length; i++) {
			gc=(GregorianCalendar)last.clone();
			gc.add(Calendar.DATE, -i);
			result[msName[DAILY].length-1-i] = gc.get(Calendar.YEAR)+"/"+
					(gc.get(Calendar.MONTH)+1)+"/"+
					gc.get(Calendar.DAY_OF_MONTH);
		}
		return result;
	}
	
	public String[] getMonthlyLabels() {
		if (!monthlyDone)
			createMonthlyData();
		
		String result[] = new String[msName[MONTHLY].length];
		GregorianCalendar gc;
		
		for (int i=0; i<msName[MONTHLY].length; i++) {
			gc=(GregorianCalendar)last.clone();
			gc.add(Calendar.MONTH, -i);
			result[msName[MONTHLY].length-1-i] = gc.get(Calendar.YEAR)+"/"+
					(gc.get(Calendar.MONTH)+1);
		}
		return result;
	}
	

	

	
	public String getHourMS(int hour) {
		if (hour<0 || hour>=msName[HOURLY].length || msName[HOURLY][hour]==null)
			return null;
		return msName[HOURLY][hour];
	}
	
	public String getHourPool(int hour) {
		if (hour<0 || hour>=poolName[HOURLY].length || poolName[HOURLY][hour]==null)
			return null;
		return poolName[HOURLY][hour];
	}
	
	public boolean getHourCap(int hour) {
		if (hour<0 || hour>=cap[HOURLY].length || msName[HOURLY][hour]==null)
			return true;
		return cap[HOURLY][hour];
	}
	
	
	public String getDayMS(int day) {
		if (day<0 || day>=msName[DAILY].length || msName[DAILY][day]==null)
			return null;
		return msName[DAILY][day];
	}
	
	public String getDayPool(int day) {
		if (day<0 || day>=poolName[DAILY].length || poolName[DAILY][day]==null)
			return null;
		return poolName[DAILY][day];
	}
	
	public boolean getDayCap(int day) {
		if (day<0 || day>=cap[DAILY].length || msName[DAILY][day]==null)
			return true;
		return cap[DAILY][day];
	}
	
	
	public String getMonthMS(int month) {
		if (!monthlyDone)
			createMonthlyData();
		if (msName[MONTHLY]==null || month<0 || month>msName[MONTHLY].length || msName[MONTHLY][month]==null)
			return null;
		return msName[MONTHLY][month];
	}
	
	public String getMonthPool(int month) {
		if (!monthlyDone)
			createMonthlyData();
		if (poolName[MONTHLY]==null || month<0 || month>poolName[MONTHLY].length || poolName[MONTHLY][month]==null)
			return null;
		return poolName[MONTHLY][month];
	}
	
	public boolean getMonthCap(int month) {
		if (!monthlyDone)
			createMonthlyData();
		if (cap[MONTHLY]==null || month<0 || month>cap[MONTHLY].length)
			return true;
		return cap[MONTHLY][month];
	}
	
	public void sanitize(StringChanger msChange, StringChanger poolChange ) {
		int i,j;
		
		for (i=0; msNames!=null && i<msNames.size(); i++)
			msNames.setElementAt(msChange.translate(msNames.elementAt(i)), i);
		
		for (i=0; poolNames!=null && i<poolNames.size(); i++)
			if (!poolNames.elementAt(i).equals("DefaultPool"))
				poolNames.setElementAt(poolChange.translate(poolNames.elementAt(i)), i);
		
		for (i=0; msName!=null && i<msName.length; i++)
			for (j=0; msName[i]!=null && j<msName[i].length; j++)
				msName[i][j] = msChange.translate(msName[i][j]);
		
		for (i=0; poolName!=null && i<poolName.length; i++)
			for (j=0; poolName[i]!=null && j<poolName[i].length; j++)
				if (poolName[i][j]!=null && !poolName[i][j].equals("DefaultPool"))
					poolName[i][j] = poolChange.translate(poolName[i][j]);
	}
	
}
