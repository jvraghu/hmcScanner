package hmcScanner;

import java.util.Calendar;
import java.util.GregorianCalendar;
import java.util.Vector;

public class LparStatus {
	
	private static int DAYS = 365;
	private static int HOURLY_DAYS = 60;
	
	private GregorianCalendar last;
	
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
	
	private Vector<String>	names = new Vector<String>();
	
	private Vector<String>	msNames = new Vector<String>();
	private Vector<String>	poolNames = new Vector<String>();
	
	
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
	
	
	
	public LparStatus(GregorianCalendar last) {
		this.last = last;
	}
	

	private void allocate() {
		msDaily = new String[DAYS];
		poolDaily = new String[DAYS];
		capDaily = new boolean[DAYS];

		for (int i=0; i<DAYS; i++) {
			msDaily[i] = null;
			poolDaily[i] = null;
			capDaily[i] = true;
		}
		
		msHourly = new String[HOURLY_DAYS*24];
		poolHourly = new String[HOURLY_DAYS*24];
		capHourly = new boolean[HOURLY_DAYS*24];
		
		for (int i=0; i<HOURLY_DAYS*24; i++) {
			msHourly[i]=null;
			poolHourly[i]=null;
			capHourly[i]=true;
		}
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
	
	
	private String getString(String s) {
		int i;
		
		for (i=0; i<names.size(); i++) {
			if (s.equals(names.elementAt(i)))
				break;
		}
		
		if (i<names.size())
			return names.elementAt(i);
		
		names.add(s);
		
		return s;
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
	
	
	/*
	 * Add data either hourly or daily. Daily data is discarded if hourly data is present.
	 */
	public void add(GregorianCalendar gc, String ms, String pool, boolean cap) {
		int slot;
		
		if (ms==null || pool==null)
			return;
		
		if (pool.equals(""))
			pool = "DefaultPool";
		
		if (msDaily==null) {
			allocate();
		}
			
		if (last==null) {
			last = gc;
		}
		
		if (status==FINISHED)
			return;
		

		slot = getSlot(gc);
		if (slot >=0 && slot < DAYS && msDaily[slot]==null) {
			msDaily[slot] = getMSString(ms);
			poolDaily[slot] = getPoolString(pool);
			capDaily[slot] = cap;
		}


		
		// Setup hourly data	
		if (status==HOURLY) {
			slot = getHourSlot(gc);
			if (slot >=0 && slot < HOURLY_DAYS*24 && msHourly[slot]==null) {
				msHourly[slot] = getMSString(ms);
				poolHourly[slot] = getPoolString(pool);
				capHourly[slot] = cap;
			}
		}	
		
	}
	
	public void endOfHourlyData() {
		status = DAILY;
	}
	
	
	/*
	 * End of data input. Compute weekly and monthly data
	 */
	public void endOfDailyData() {
		if (msDaily==null)
			return;
		
		msMonthly = new String[NUM_MONTHS];
		poolMonthly = new String[NUM_MONTHS];
		capMonthly = new boolean[NUM_MONTHS];
		
		monthName = new GregorianCalendar[NUM_MONTHS];
		
		for (int i=0; i<NUM_MONTHS; i++) {
			msMonthly[i] = null;
			poolMonthly[i] = null;
			capMonthly[i] = true;
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
			} else if (gc.get(Calendar.MONTH)!=monthName[m].get(Calendar.MONTH)) {
				// Month is changed
				m++;
				monthName[m]=gc;
			}
			
			if (msDaily[d]!=null) {
				msMonthly[m]=msDaily[d];
				poolMonthly[m]=poolDaily[d];
				capMonthly[m]=capDaily[d];
			}
		}
	
		status=FINISHED;		
	}
	
	
	// hour=0 is the first slot
	public GregorianCalendar getHourDate(int hour) {
		if (status != FINISHED || hour<0 || hour>=HOURLY_DAYS*24)
			return null;
		
		GregorianCalendar gc = (GregorianCalendar)last.clone();
		gc.add(Calendar.HOUR, -(HOURLY_DAYS*24-1-hour) );
		return gc;
	}
	
	public String getHourMS(int hour) {
		if (status != FINISHED || hour<0 || hour>=HOURLY_DAYS*24 || msHourly[hour]==null)
			return null;
		return msHourly[hour];
	}
	
	public String getHourPool(int hour) {
		if (status != FINISHED || hour<0 || hour>=HOURLY_DAYS*24 || msHourly[hour]==null)
			return null;
		return poolHourly[hour];
	}
	
	public boolean getHourCap(int hour) {
		if (status != FINISHED || hour<0 || hour>=HOURLY_DAYS*24 || msHourly[hour]==null)
			return true;
		return capHourly[hour];
	}
	
	
	public GregorianCalendar getDayDate(int day) {
		if (status != FINISHED || day<0 || day>=DAYS)
			return null;
		
		GregorianCalendar gc = (GregorianCalendar)last.clone();
		gc.add(Calendar.DAY_OF_MONTH, -(DAYS-1-day) );
		return gc;
	}
	
	public String getDayMS(int day) {
		if (status != FINISHED || day<0 || day>=DAYS || msDaily[day]==null)
			return null;
		return msDaily[day];
	}
	
	public String getDayPool(int day) {
		if (status != FINISHED || day<0 || day>=DAYS || msDaily[day]==null)
			return null;
		return poolDaily[day];
	}
	
	public boolean getDayCap(int day) {
		if (status != FINISHED || day<0 || day>=DAYS || msDaily[day]==null)
			return true;
		return capDaily[day];
	}
	
	
	public GregorianCalendar[] getMonthNames() {
		if (status != FINISHED)
			return null;
		return monthName;
	}
	
	public String getMonthMS(int month) {
		if (status != FINISHED || msMonthly[month]==null)
			return null;
		return msMonthly[month];
	}
	
	public String getMonthPool(int month) {
		if (status != FINISHED || msMonthly[month]==null)
			return null;
		return poolMonthly[month];
	}
	
	public boolean getMonthCap(int month) {
		if (status != FINISHED || msMonthly[month]==null)
			return true;
		return capMonthly[month];
	}
	
	public void sanitize(StringChanger msChange, StringChanger poolChange ) {
		int i;
		
		for (i=0; msNames!=null && i<msNames.size(); i++)
			msNames.setElementAt(msChange.translate(msNames.elementAt(i)), i);
		
		for (i=0; poolNames!=null && i<poolNames.size(); i++)
			if (!poolNames.elementAt(i).equals("DefaultPool"))
				poolNames.setElementAt(poolChange.translate(poolNames.elementAt(i)), i);		
		
		for (i=0; msDaily!=null && i<msDaily.length; i++)
			if (msDaily[i]!=null)
				msDaily[i]=msChange.translate(msDaily[i]);
		for (i=0; msHourly!=null && i<msHourly.length; i++)
			if (msHourly[i]!=null)
				msHourly[i]=msChange.translate(msHourly[i]);
		for (i=0; msMonthly!=null && i<msMonthly.length; i++)
			if (msMonthly[i]!=null)
				msMonthly[i]=msChange.translate(msMonthly[i]);
		
		for (i=0; poolDaily!=null && i<poolDaily.length; i++) {
			if (poolDaily[i]!=null && !poolDaily[i].equals("DefaultPool"))
				poolDaily[i]=poolChange.translate(poolDaily[i]);
		}
		for (i=0; poolHourly!=null && i<poolHourly.length; i++) {
			if (poolHourly[i]!=null && !poolHourly[i].equals("DefaultPool"))
				poolHourly[i]=poolChange.translate(poolHourly[i]);
		}
		for (i=0; poolMonthly!=null && i<poolMonthly.length; i++) {
			if (poolMonthly[i]!=null && !poolMonthly[i].equals("DefaultPool"))
				poolMonthly[i]=poolChange.translate(poolMonthly[i]);
		}
	}
	
}
