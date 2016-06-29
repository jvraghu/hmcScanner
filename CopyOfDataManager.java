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

public class CopyOfDataManager {
	
	private static int DAYS = 365;
	private static int HOURLY_DAYS = 60;
	private int NUM_COUNTERS = 100;				// level counters for value distribution
	private static float MAX_ERROR = 0.10f;		// error in 90-95 finding
	
	private GregorianCalendar last;
	
	// Daily data
	private float[]		min, avg, max;		// Data related to a fully year, day by day
	private int[]		num;
	
	
	private static byte	HOURLY	= 0;
	private static byte DAILY	= 1;
	private static byte FINISHED = 2;
	private byte status = HOURLY;
	
	private static byte MIN_HOURLY_SAMPLES = 0;   // min 3/4 of 24 hour data
	
	// Monthly data
	private float[]		monthMin, monthAvg, monthMax;		// full year, month by month
	private float[]		monthNum;
	private GregorianCalendar[]	monthName;
	
	private static byte NUM_MONTHS = 13;		// 13 months due to start in middle of a month!!
	
	// Hourly data
	private float[]		hourAvg;		
	private float[]		hourNum;
	
	// Levels
	private double		daily90p 	= 0; 
	private double		daily95p 	= 0;
	private double		hourly90p 	= 0;
	private double		hourly95p 	= 0;
	private double		dailyMax 	= 0;
	private double		dailyAvg 	= 0;
	private double		hourlyMax 	= 0;
	private double		hourlyAvg 	= 0;
	
	
	// StatManagers
	private StatManager dailyStat = new StatManager();
	private StatManager hourlyStat = new StatManager();
	
	
	
	
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
	

	public CopyOfDataManager() {
		last = null;
		min=avg=max=null;
		num=null;
	}
	
	
	public CopyOfDataManager(GregorianCalendar last) {
		this.last = last;
		min=avg=max=null;
		num=null;
	}
	
	
	private void allocate() {
		min = new float[DAYS];
		max = new float[DAYS];
		avg = new float[DAYS];
		num = new int[DAYS];
		
		for (int i=0; i<DAYS; i++) {
			min[i]=Float.POSITIVE_INFINITY;
			max[i]=Float.NEGATIVE_INFINITY;
			avg[i]=0;
			num[i]=0;
		}
		
		hourAvg = new float[HOURLY_DAYS*24];
		hourNum = new float[HOURLY_DAYS*24];
		
		for (int i=0; i<HOURLY_DAYS*24; i++) {
			hourAvg[i]=0;
			hourNum[i]=0;
		}
		
		dailyStat.setStep(.01);
		hourlyStat.setStep(.01);
	}
	
	private int getSlot(GregorianCalendar gc) {
		float diff;
		int day;
		
		diff = last.getTimeInMillis()-gc.getTimeInMillis();
		day = (DAYS-1) - (int)(diff/(24*60*60*1000));
		
		return day;
	}
	
	
	private int getHourSlot(GregorianCalendar gc) {
		float diff;
		int slot;
		
		diff = last.getTimeInMillis()-gc.getTimeInMillis();
		slot = (HOURLY_DAYS*24-1) - (int)(diff/(60*60*1000));

		return slot;
	}
	
	
	/*
	 * Add data either hourly or daily. Daily data is discarded if hourly data is present.
	 */
	public void add(GregorianCalendar gc, float value) {
		int slot;
		
		if (num==null) {
			allocate();
		}
			
		if (last==null) {
			last = gc;
		}
		
		if (value<0)
			return;
		
		slot = getSlot(gc);
		if (slot <0 || slot >= DAYS)
			return;
		
		if (status==DAILY)
			dailyStat.addNumber(value);
		
		// If status is DAILY we only accept data in empty slots
		if (status==FINISHED || (status == DAILY && num[slot]!=0) )
			return;
		
		// Setup daily data
		if (value < min[slot]) min[slot]=value;
		if (value > max[slot]) max[slot]=value;
		
		avg[slot] += value;
		num[slot]++;	
		
		
		
		// Setup hourly data	
		if (status!=HOURLY)
			return;
		slot = getHourSlot(gc);
		if (slot <0 || slot >= HOURLY_DAYS*24)
			return;
		hourAvg[slot] += value;
		hourNum[slot]++;	
		
		hourlyStat.addNumber(value);
	}
	
	/*
	 * Remove days with limited data included and change status
	 */
	public void endOfHourlyData() {
		if (num==null)
			return;
		
		for (int i=0; i<DAYS; i++) {
			if (num[i]<MIN_HOURLY_SAMPLES) {
				min[i]=Float.POSITIVE_INFINITY;
				max[i]=Float.NEGATIVE_INFINITY;
				avg[i]=0;
				num[i]=0;
				continue;
			}
			if (num[i]!=0)
				avg[i] = avg[i] / num[i];
		}
		
		
		for (int i=0; i<HOURLY_DAYS*24; i++) {
			if (hourNum[i]!=0)
				hourAvg[i] = hourAvg[i] / hourNum[i];
		}
		
		
		status = DAILY;
	}
	
	/*
	 * End of data input. Compute weekly and monthly data
	 */
	public void endOfDailyData() {
		if (num==null)
			return;
		
		monthAvg = new float[NUM_MONTHS];
		monthMin = new float[NUM_MONTHS];
		monthMax = new float[NUM_MONTHS];
		monthNum = new float[NUM_MONTHS];
		monthName = new GregorianCalendar[NUM_MONTHS];
		
		for (int i=0; i<NUM_MONTHS; i++) {
			monthAvg[i] = 0;
			monthMin[i] = Float.POSITIVE_INFINITY;
			monthMax[i] = Float.NEGATIVE_INFINITY;
			monthNum[i] = 0;
			monthName[i] = null;
		}
		
		int m = 0;					// month index
		int d;						// day index
		GregorianCalendar gc;		// current day
		
		for (d=0; d<DAYS; d++) {
			gc = (GregorianCalendar)last.clone();
			gc.add(Calendar.DAY_OF_MONTH, -(DAYS-1-d) );
			
			if (monthName[m]==null) {
				// first sample
				monthName[m]=gc;
				monthNum[m]=0;
			} else if (gc.get(Calendar.MONTH)!=monthName[m].get(Calendar.MONTH)) {
				// Month is changed
				m++;
				monthName[m]=gc;
				monthNum[m]=0;
			}
			
			if (num[d]>0) {
				monthAvg[m] += avg[d];
				if (min[d] < monthMin[m]) monthMin[m]=min[d];
				if (max[d] > monthMax[m]) monthMax[m]=max[d];
				monthNum[m]++;	
			}
		}
		
		boolean missing_month=false;
		int good_month=-1;
		for (m=0; m<NUM_MONTHS; m++) {
			if (monthNum[m]!=0) {
				monthAvg[m] = monthAvg[m] / monthNum[m];
				good_month=m;
			} else
				missing_month=true;
		}
		
		if (missing_month && good_month>=0) {
			for (m=0; m<NUM_MONTHS; m++) {
				if (monthName[m]==null) {
					gc = (GregorianCalendar)monthName[good_month].clone();
					gc.add(Calendar.MONTH, m-good_month);
					monthName[m]=gc;
				}
			}
		}
		
		status=FINISHED;	
		
		// Compute levels
		int i,j,n;
		int[]		level;		// counters for level distribution
		int			numSamples = 0;
		double		delta;

		
		
		// DAILY
		dailyMax=0;
		dailyAvg=0;
		
		for (i=0; i<avg.length; i++) {
			if (num[i]>0) {
				numSamples++;
				dailyAvg+=avg[i];
				if (avg[i]>dailyMax)
					dailyMax=avg[i];
			}
		}
		
		if (numSamples==0) 
			return;
		
		dailyAvg = dailyAvg / numSamples;
		
		// No not compure percentiles if minimal delta
		if (dailyMax-dailyAvg>MAX_ERROR) {	
			
			NUM_COUNTERS = 1 + (int)( (dailyMax-dailyAvg) / MAX_ERROR );
			
			delta = (dailyMax-dailyAvg)/NUM_COUNTERS;
			level = new int[NUM_COUNTERS];
			
			for (i=0; i<avg.length; i++) {
				if (num[i]>0) {
					if (avg[i]<=dailyAvg)
						n = 0;
					else
						n = (int)(1d*(avg[i]-dailyAvg)/delta)+1;
					for (j=n; j<NUM_COUNTERS; j++)
						level[j]++;
				}
			}
			
			n=0;
			while ( n<NUM_COUNTERS && level[n]*100d<=numSamples*90d )
				n++;
			daily90p = dailyAvg + delta * n;
			if (daily90p>dailyMax)
				daily90p=dailyMax;
			
			n=0;
			while ( n<NUM_COUNTERS && level[n]*100d<=numSamples*95d )
				n++;
			daily95p = dailyAvg + delta * n;
			if (daily95p>dailyMax)
				daily95p=dailyMax;
		} else {
			daily90p = daily95p = dailyMax;
		}
		
		
		// HOURLY
		hourlyMax=0;
		hourlyAvg=0;
		
		numSamples=0;
		for (i=0; i<hourAvg.length; i++) {
			if (hourNum[i]>0) {
				numSamples++;
				hourlyAvg+=hourAvg[i];
				if (hourAvg[i]>hourlyMax)
					hourlyMax=hourAvg[i];
			}
		}
		
		if (numSamples==0) 
			return;
		
		hourlyAvg = hourlyAvg / numSamples;
		
		// No not compure percentiles if minimal delta
		if (hourlyMax-hourlyAvg>MAX_ERROR) {	
			
			NUM_COUNTERS = 1 + (int)( (hourlyMax-hourlyAvg) / MAX_ERROR );
		
			delta = (hourlyMax-hourlyAvg)/NUM_COUNTERS;
			level = new int[NUM_COUNTERS];
			
			for (i=0; i<hourAvg.length; i++) {
				if (hourNum[i]>0) {
					if (hourAvg[i]<=hourlyAvg)
						n = 0;
					else
						n = (int)((1d*(hourAvg[i]-hourlyAvg))/delta)+1;
					for (j=n; j<NUM_COUNTERS; j++)
						level[j]++;
				}
			}
			
			n=0;
			while ( n<NUM_COUNTERS && level[n]*100d<=numSamples*90d )
				n++;
			hourly90p = hourlyAvg + delta * n;
			if (hourly90p>hourlyMax) 
				hourly90p=hourlyMax;
			
			n=0;
			while ( n<NUM_COUNTERS && level[n]*100d<=numSamples*95d )
				n++;
			hourly95p = hourlyAvg + delta * n;
			if (hourly95p>hourlyMax) 
				hourly95p=hourlyMax;
		} else {
			hourly95p = hourly90p = hourlyMax;
		}
		
	}
	
	

	
	
	
	
	
	// hour=0 is the first slot
	public GregorianCalendar getHourDate(int hour) {
		if (status != FINISHED || hour<0 || hour>=HOURLY_DAYS*24)
			return null;
		
		GregorianCalendar gc = (GregorianCalendar)last.clone();
		gc.add(Calendar.HOUR, -(HOURLY_DAYS*24-1-hour) );
		return gc;
	}
	
	public float getHourAvg(int hour) {
		if (status != FINISHED || hour<0 || hour>=HOURLY_DAYS*24 || hourNum[hour]==0)
			return -1;
		return hourAvg[hour];
	}
	
	
	public GregorianCalendar getDayDate(int day) {
		if (status != FINISHED || day<0 || day>=DAYS)
			return null;
		
		GregorianCalendar gc = (GregorianCalendar)last.clone();
		gc.add(Calendar.DAY_OF_MONTH, -(DAYS-1-day) );
		return gc;
	}
	
	public float getDayAvg(int day) {
		if (status != FINISHED || day<0 || day>=DAYS || num[day]==0)
			return -1;
		return avg[day];
	}
	
	public float getDayMin(int day) {
		if (status != FINISHED || day<0 || day>=DAYS || num[day]==0)
			return -1;
		return min[day];
	}
	
	public float getDayMax(int day) {
		if (status != FINISHED || day<0 || day>=DAYS || num[day]==0)
			return -1;
		return max[day];
	}
	
	
	public GregorianCalendar[] getMonthNames() {
		if (status != FINISHED)
			return null;
		return monthName;
	}
	
	public float getMonthMin(int month) {
		if (status != FINISHED || monthNum[month]==0)
			return -1;
		return monthMin[month];
	}
	
	public float getMonthMax(int month) {
		if (status != FINISHED || monthNum[month]==0)
			return -1;
		return monthMax[month];
	}
	

	public float getMonthAvg(int month) {
		if (status != FINISHED || monthNum[month]==0)
			return -1;
		return monthAvg[month];
	}
	
	public float getMonthNum(int month) {
		if (status != FINISHED)
			return -1;
		return monthNum[month];
	}
	
	public double getDaily90p() { return daily90p; }
	public double getDaily95p() { return daily95p; }
	public double getHourly90p() { 
		double d = hourlyStat.getLevel(.9);
		System.out.println("object="+d+"\tinternal="+hourly90p);
		return hourly90p; 
	}
	
	public double getHourly95p() { 
		double d = hourlyStat.getLevel(.95);
		//System.out.println("object="+d+"\tinternal="+hourly95p);
		return hourly95p; 
	}
	public double
	getDailyMax() { return dailyMax; }
	public double getDailyAvg() { return dailyAvg; }
	public double getHourlyMax() { return hourlyMax; }
	public double getHourlyAvg() { return hourlyAvg; }
}
