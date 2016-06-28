package hmcScanner;

import java.util.Calendar;
import java.util.GregorianCalendar;

public class DataManager {
	
	private static int DAYS = 365;
	private static int HOURLY_DAYS = 60;

	private GregorianCalendar last;
	
	// data
	private float[]		dailyData = new float[365];
	private float[]		hourlyData= new float[24*30*2];
	private float[]		monthlyData = null;
	

	// StatManagers
	private StatManager dailyStat = new StatManager();
	private StatManager hourlyStat = new StatManager();
	private StatManager monthlyStat[] = new StatManager[13];


	
	
	private void createMonthlyData() {
		float	sum = 0;
		float	num = 0;
		int 	month, day;
		GregorianCalendar curr, prev;
		
		for (month=0; month<13; month++)
			monthlyStat[month]=new StatManager();
		
		monthlyData = new float[13];
		curr = (GregorianCalendar)last.clone();
		prev = (GregorianCalendar)last.clone();
		
		for (day=dailyData.length-1, month=12; day>=0; day--) {
			if (curr.get(Calendar.MONTH)==prev.get(Calendar.MONTH)) {
				if (dailyData[day]>=0) {
					sum += dailyData[day];
					num++;
					monthlyStat[month].addNumber(dailyData[day]);
				}
			} else {
				if (num>0)
					monthlyData[month]=sum/num;
				else
					monthlyData[month]=-1;
				month--;
				if (dailyData[day]>=0) {
					sum = dailyData[day];
					num = 1;
					monthlyStat[month].addNumber(dailyData[day]);
				} else {
					sum = 0;
					num = 0;
				}
				prev = (GregorianCalendar)curr.clone();
			}
			curr.add(Calendar.DATE, -1);			
		}
		
		if (num>0)
			monthlyData[month]=sum/num;
		else
			monthlyData[month]=-1;		
	}
	
	
	
	public void addHourData(GregorianCalendar gc, float value) {
		int slot = getHourSlot(gc);
		if (slot<0 || slot>=hourlyData.length) {
			/*
			System.out.println("Out of bound hourly data insert: "+gc.get(Calendar.YEAR)+"/"+
														(gc.get(Calendar.MONTH)+1)+"/"+
														gc.get(Calendar.DAY_OF_MONTH)+" "+
														gc.get(Calendar.HOUR_OF_DAY)+":"+
														gc.get(Calendar.MINUTE)+"/"+
														gc.get(Calendar.SECOND) );
			*/
			return;
		}
		if (hourlyData[slot]>=0) {
			/*
			System.out.println("Warning! Two samples added in same hour: "+gc.get(Calendar.YEAR)+"/"+
														(gc.get(Calendar.MONTH)+1)+"/"+
														gc.get(Calendar.DAY_OF_MONTH)+" "+
														gc.get(Calendar.HOUR_OF_DAY)+":"+
														gc.get(Calendar.MINUTE)+"/"+
														gc.get(Calendar.SECOND) );
			*/
			;
		}
		
		hourlyData[slot]=value;
		
		hourlyStat.addNumber(value);
	}
	
	public float getHourData(GregorianCalendar gc) {
		int slot = getHourSlot(gc);
		if (slot<0 || slot>=hourlyData.length) {
			System.out.println("Out of bound hourly data request: "+gc.get(Calendar.YEAR)+"/"+
														(gc.get(Calendar.MONTH)+1)+"/"+
														gc.get(Calendar.DAY_OF_MONTH)+" "+
														gc.get(Calendar.HOUR_OF_DAY)+":"+
														gc.get(Calendar.MINUTE)+"/"+
														gc.get(Calendar.SECOND) );
			return -1;
		}
		
		return hourlyData[slot];
	}
	
	public float getHourData(int slot) {
		if (slot<0 || slot>=hourlyData.length) {
			System.out.println("Out of bound hourly data request: slot " + slot );
			return -1;
		}
		
		return hourlyData[slot];
	}
	
	public float getMonthData(int slot) {
		if (monthlyData==null)
			createMonthlyData();
		
		if (slot<0 || slot>=monthlyData.length) {
			System.out.println("Out of bound monthly data request: slot " + slot );
			return -1;
		}
		
		return monthlyData[slot];
	}
	
	public void addDayData(GregorianCalendar gc, float value) {
		int slot = getSlot(gc);
		if (slot<0 || slot>=dailyData.length) {
			/*
			System.out.println("Out of bound daily data insert: "+gc.get(Calendar.YEAR)+"/"+
														(gc.get(Calendar.MONTH)+1)+"/"+
														gc.get(Calendar.DAY_OF_MONTH)+" "+
														gc.get(Calendar.HOUR_OF_DAY)+":"+
														gc.get(Calendar.MINUTE)+":"+
														gc.get(Calendar.SECOND) );
			*/
			return;
		}
		if (dailyData[slot]>=0) {
			/*
			System.out.println("Warning! Two samples added in same day: "+gc.get(Calendar.YEAR)+"/"+
														(gc.get(Calendar.MONTH)+1)+"/"+
														gc.get(Calendar.DAY_OF_MONTH)+" "+
														gc.get(Calendar.HOUR_OF_DAY)+":"+
														gc.get(Calendar.MINUTE)+":"+
														gc.get(Calendar.SECOND) );
			*/
			;
		}
		
		dailyData[slot]=value;
		
		dailyStat.addNumber(value);
	}
	
	public float getDayData(GregorianCalendar gc) {
		int slot = getSlot(gc);
		if (slot<0 || slot>=dailyData.length) {
			System.out.println("Out of bound daily data request: "+gc.get(Calendar.YEAR)+"/"+
														(gc.get(Calendar.MONTH)+1)+"/"+
														gc.get(Calendar.DAY_OF_MONTH)+" "+
														gc.get(Calendar.HOUR_OF_DAY)+":"+
														gc.get(Calendar.MINUTE)+"/"+
														gc.get(Calendar.SECOND) );
			return -1;
		}
		
		return dailyData[slot];
	}
	
	public float getDayData(int slot) {
		if (slot<0 || slot>=dailyData.length) {
			System.out.println("Out of bound daily data request: slot" + slot );
			return -1;
		}
		
		return dailyData[slot];
	}
	
	
	public DataManager(GregorianCalendar last) {
		
		// Set it to 23:59 in order to get the entire day
		this.last = (GregorianCalendar)last.clone();
		this.last.set(Calendar.HOUR_OF_DAY, 23);
		this.last.set(Calendar.MINUTE, 59);

		// Reset data
		int i;
		for (i=0; i<hourlyData.length; i++)
			hourlyData[i]=-1;
		for (i=0; i<dailyData.length; i++)
			dailyData[i]=-1;
		
		dailyStat.setStep(.01);
		hourlyStat.setStep(.01);
	}
	
	
	public String[] getHourlyLabels() {
		String result[] = new String[hourlyData.length];
		GregorianCalendar gc;
		
		for (int i=0; i<hourlyData.length; i++) {
			gc=(GregorianCalendar)last.clone();
			gc.add(Calendar.HOUR, -i);
			result[hourlyData.length-1-i] = gc.get(Calendar.YEAR)+"/"+
					(gc.get(Calendar.MONTH)+1)+"/"+
					gc.get(Calendar.DAY_OF_MONTH)+" "+
					gc.get(Calendar.HOUR_OF_DAY);
		}
		return result;
	}
	
	public String getHourLabel(int n) {
		GregorianCalendar gc;
		gc=(GregorianCalendar)last.clone();
		gc.add(Calendar.HOUR, n-hourlyData.length+1);
		return gc.get(Calendar.YEAR)+"/"+
				(gc.get(Calendar.MONTH)+1)+"/"+
				gc.get(Calendar.DAY_OF_MONTH)+" "+
				gc.get(Calendar.HOUR_OF_DAY);
	}
	
	public String[] getDailyLabels() {
		String result[] = new String[dailyData.length];
		GregorianCalendar gc;
		
		for (int i=0; i<dailyData.length; i++) {
			gc=(GregorianCalendar)last.clone();
			gc.add(Calendar.DATE, -i);
			result[dailyData.length-1-i] = gc.get(Calendar.YEAR)+"/"+
					(gc.get(Calendar.MONTH)+1)+"/"+
					gc.get(Calendar.DAY_OF_MONTH);
		}
		return result;
	}
	
	public String getDayLabel(int n) {
		GregorianCalendar gc;
		gc=(GregorianCalendar)last.clone();
		gc.add(Calendar.HOUR, n-dailyData.length+1);
		return gc.get(Calendar.YEAR)+"/"+
				(gc.get(Calendar.MONTH)+1)+"/"+
				gc.get(Calendar.DAY_OF_MONTH);
	}
	
	
	
	public String[] getMonthlyLabels() {
		if (monthlyData==null)
			createMonthlyData();
		
		String result[] = new String[monthlyData.length];
		GregorianCalendar gc;
		
		for (int i=0; i<monthlyData.length; i++) {
			gc=(GregorianCalendar)last.clone();
			gc.add(Calendar.MONTH, -i);
			result[monthlyData.length-1-i] = gc.get(Calendar.YEAR)+"/"+
					(gc.get(Calendar.MONTH)+1);
		}
		return result;
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
	
	

	public double getDaily90p() { return dailyStat.getLevel(.9); }
	public double getDaily95p() { return dailyStat.getLevel(.95); }
	
	public double getHourly90p() { return hourlyStat.getLevel(.9);	}	
	public double getHourly95p() { return hourlyStat.getLevel(.95);	}
	
	public double getDailyMax() { return dailyStat.getMax(); }
	public double getDailyAvg() { return dailyStat.getAvg(); }
	
	public double getHourlyMax() { return hourlyStat.getMax(); }
	public double getHourlyAvg() { return hourlyStat.getAvg(); }
	
	public double getMonthlyMax(int i) { 
		if (monthlyData==null)
			createMonthlyData();
		return monthlyStat[i].getMax(); 
	}
	
	public double getMonthlyAvg(int i) { 
		if (monthlyData==null)
			createMonthlyData();
		return monthlyStat[i].getAvg(); 
	}
}
