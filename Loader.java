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

import java.awt.Graphics2D;
import java.awt.image.BufferedImage;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.Console;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.PrintWriter;
import java.math.BigDecimal;
import java.math.MathContext;
import java.util.Calendar;
import java.util.GregorianCalendar;
import java.util.Locale;
import java.util.Vector;
import java.util.zip.GZIPInputStream;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.format.Orientation;
import jxl.format.VerticalAlignment;
import jxl.read.biff.BiffException;
import jxl.write.Formula;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.NumberFormat;
import jxl.write.NumberFormats;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableImage;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class Loader {
	
	private static byte		NUM_RETRY		= 3;
	
	// DEBUG ONLY!!! Always set to false
	private boolean 	onlyReadFile = true;
	private boolean		produceStatistics = false;
	private boolean		rowMode = false;
	private boolean		verboseMode = false;
	
	private static String	version = "0.11.34";
	
	private String		_hmc = null;
	private String		user = null;
	private String		password = null;
	private String		_baseDir = null;
	private String		startPerf = null;
	private String		endPerf = null;
	private String		samplePerf = null;
	private SSHManager2	sshm = null;
	private String		sshLogfile = null;
	private String		sshPrivKey = null;
	private String		proxyType = null;
	private String		proxyHost = null;
	private String		proxyPort = null;
	private String		proxyUser = null;
	private String		proxyPassword = null;
	private String		selectedMS = null;
	private char		csvSeparator = ',';
	private int			timeout = 0;
	private boolean 	novios = false;
	
	
	private GenericData	managedSystem[] = null;
	private GenericData	scannerParams;
	private GregorianCalendar scannerDate=null;
	private GenericData hmc = null;
	private GenericData entPool[] = null;
	
	private byte		managerType = M_UNKNOWN;
	
	private WritableWorkbook workbook = null;
	
	private static final BigDecimal zero = new BigDecimal(0);
	
	// File names
	private static String	excel = "_scan.xls";
	private static String	excelSane = "_scan_safe.xls";
	private static String	systemData = "system_data.txt"; 
	private static String	procSysData = "system_proc.txt";
	private static String	memSysData = "system_mem.txt";
	private static String	slotSysData = "system_slot.txt";
	private static String	procLparData = "lpar_proc.txt";
	private static String	memLparData = "mem_proc.txt";
	private static String	procPoolData = "procpool.txt";
	private static String	memPoolData = "mempool.txt";
	private static String	lparConfigData = "lpar_config.txt";
	private static String	vswitchData = "vswitch.txt";
	private static String	vethData = "veth.txt";
	private static String	vscsiData = "vscsi.txt";
	private static String	vfcData = "vfc.txt";
	private static String	utilDataConfig = "util_data_config.txt";
	private static String	lslparutilData = "lslparutil.txt.gz";
	private static String	viosList = "vioslist.txt";
	private static String	npivData = "npiv_data.txt";
	private static String	vscsiDiskData = "vscsidisk.txt";
	private static String	lshmcv = "lshmc-v.txt";
	private static String	lshmcn = "lshmc-n.txt";
	private static String	lshmcb = "lshmc-b.txt";
	//private static String	lshmcr = "lshmc-r.txt";
	private static String	lshmcV = "lshmc-vv.txt";
	private static String	scannerInfo = "scannerInfo.txt";
	private static String	lastPerfDate = "lastPerfData.txt";
	private static String	systemPerf = "systemPerf.txt.gz";
	private static String	lparPerf = "lparPerf.txt";
	private static String	lslparutilStats = "lslparutilStat.txt.gz";
	private static String	lslicSyspower = "lslic_syspower.txt";
	private static String	diskuuid = "disk_uuid.txt";
	private static String	hmcdate = "hmc-date.txt";
	private static String	ioslevel = "ioslevel.txt";
	private static String	lspartition = "lspartition.txt";
	private static String	lslparutilHourlyStats = "lslparutilHourlyStat.txt.gz";
	private static String	lslparutilDailyStats = "lslparutilDailyStat.txt.gz";
	private static String	hmcScannerPic="hmcScanner.png";
	private static String	ivmversion="ivmversion.txt";
	private static String	seaCfg="seacfg.txt";
	private static String	slots="slots.txt";
	private static String	etherChannel="etherChannel.txt";
	private static String	entstatSEA="entstatSEA.txt";
	private static String	fcstat="fcstat.txt";
	private static String	fcattr="fcattr.txt";
	private static String	html="_HTML";
	private static String	html_sane="_HTML_SANE";
	private static String	lparProfiles="profiles.txt";
	private static String	lspv_size="lspv_size.txt";
	private static String	lspv_free="lspv_free.txt";
	private static String	slotchildren="lshwres-slotchildren.txt";
	private static String 	csv="_CSV";
	private static String	csv_sane="_CSV_SANE";
	private static String	lscod_bill_proc="lscod_bill_proc.txt";
	private static String	lscod_bill_mem="lscod_bill_mem.txt";
	private static String	lscod_cap_proc_onoff="lscod_cap_proc_onoff.txt";
	private static String	lscod_cap_mem_onoff="lscod_cap_mem_onoff.txt";
	private static String	lscod_hist="lscod_hist.txt";
	private static String	proc0="proc0.txt";
	private static String	lscodpool="lscodpool.txt";
	
	// HTML file names
	private static String	header_html="header.html";
	private static String	hmc_html="hmc.html";
	private static String	systems_html="systems.html";
	private static String	slots_html="slots.html";
	private static String	index_html="index.html";
	private static String	menu_html="fedmenu.html";
	private static String	lpar_html="lpar.html";
	private static String	cpu_html="cpu.html";
	private static String	profile_html="profile.html";
	private static String	mem_html="mem.html";
	private static String	iochildren_html="iochildren.html";
	private static String	veth_html="veth.html";
	private static String	vscsi_html="vscsi.html";
	private static String	vscsimap_html="vscsimap.html";
	private static String	vfc_html="vfc.html";
	private static String	viosdisks_html="viosdisks.html";
	private static String	sea_html="sea.html";
	private static String	pfc_html="pfc.html";
	private static String	poolcpu_html="poolcpu.html";
	private static String	sysram_html="sysram.html";
	private static String	lparcpu_html="lparcpu.html";
	private static String	pooldaily_html="pooldaily.html";
	private static String	lpardaily_html="lpardaily.html";
	private static String	poolhourly_html="poolhourly.html";
	private static String	lparhourly_html="lparhourly.html";
	private static String	sysperfindex_html="sysperfindex.html";
	private static String	sysperfmenu_html="sysperfmenu.html";
	private static String	lparperfmenu_html="lparperfmenu.html";
	private static String	lparperfindex_html="lparperfindex.html";
	private static String	poolperfmenu_html="poolperfmenu.html";
	private static String	poolperfindex_html="poolperfindex.html";
	private static String	onoff_html="onoff.html";
	private static String	codlog_html="codlog.html";
	private static String	syspool_html="syspool.html";
	
	// HTML oriented structures
	private Vector<String>	buttonName = null;
	private Vector<String>	htmlName = null;
	private Vector<String>	sysButtonName = null;
	private Vector<String>	sysHtmlName = null;
	private Vector<String>	lparButtonName = null;
	private Vector<String>	lparHtmlName = null;
	private Vector<String>	poolButtonName = null;
	private Vector<String>	poolHtmlName = null;
	
	// CSV file names
	private static String	header_csv="header.csv";
	private static String	hmc_csv="hmc.csv";
	private static String	systems_csv="systems.csv";
	private static String	slots_csv="slots.csv";
	private static String	index_csv="index.csv";
	private static String	menu_csv="fedmenu.csv";
	private static String	lpar_csv="lpar.csv";
	private static String	cpu_csv="cpu.csv";
	private static String	profile_csv="profile.csv";
	private static String	mem_csv="mem.csv";
	private static String	iochildren_csv="iochildren.csv";
	private static String	veth_csv="veth.csv";
	private static String	vscsi_csv="vscsi.csv";
	private static String	vscsimap_csv="vscsimap.csv";
	private static String	vfc_csv="vfc.csv";
	private static String	viosdisks_csv="viosdisks.csv";
	private static String	sea_csv="sea.csv";
	private static String	pfc_csv="pfc.csv";
	private static String	poolcpu_csv="poolcpu.csv";
	private static String	sysram_csv="sysram.csv";
	private static String	lparcpu_csv="lparcpu.csv";
	private static String	pooldaily_csv="pooldaily.csv";
	private static String	lpardaily_csv="lpardaily.csv";
	private static String	poolhourly_csv="poolhourly.csv";
	private static String	lparhourly_csv="lparhourly.csv";
	private static String	sysperfindex_csv="sysperfindex.csv";
	private static String	sysperfmenu_csv="sysperfmenu.csv";
	private static String	lparperfmenu_csv="lparperfmenu.csv";
	private static String	lparperfindex_csv="lparperfindex.csv";
	private static String	poolperfmenu_csv="poolperfmenu.csv";
	private static String	poolperfindex_csv="poolperfindex.csv";
	private static String	onoff_csv="onoff.csv";
	private static String	codlog_csv="codlog.csv";
	private static String	syspool_csv="syspool.csv";
	
	// Object types
	private static byte		PROC			= 0;
	private static byte		MEM				= 1;
	private static byte		SLOT			= 2;
	private static byte		PROC_LPAR		= 3;
	private static byte		MEM_LPAR		= 4;
	private static byte		PROC_POOL		= 5;
	private static byte		MEM_POOL		= 6;
	private static byte		CONFIG_LPAR		= 7;
	private static byte		VSWITCH			= 8;
	private static byte		VETH			= 9;
	private static byte		VSCSI			= 10;
	private static byte		VFC				= 11;
	private static byte		VFCMAP			= 12;
	private static byte		VSCSIMAP		= 13;
	private static byte		SYSPOWERLIC		= 14;
	private static byte		HDISK			= 15;
	private static byte		SEA				= 16;
	private static byte		ETH				= 17;
	private static byte		ETHERCHANNEL	= 18;
	private static byte		ENTSTATSEA		= 19;
	private static byte		FC				= 20;
	private static byte		FCSTAT			= 21;
	private static byte		PROFILES		= 22;
	private static byte		IOSLOTCHILDREN	= 23;
	private static byte		LSCOD_BILL_PROC	= 24;
	private static byte		LSCOD_BILL_MEM	= 25;
	private static byte		LSCOD_CAP_PROC_ONOFF = 26;
	private static byte		LSCOD_CAP_MEM_ONOFF = 27;
	private static byte		LSCOD_HIST 		= 28;
	private static byte		ENTPOOLSYS		= 29;
	
	// Command line data
	
	private static byte		HMC				= 0;
	private static byte		USER			= 1;
	private static byte		PASSWORD		= 2;
	private static byte		DIR				= 3;
	private static byte		STARTPERF		= 4;
	private static byte		ENDPERF			= 5;
	private static byte		SAMPLE			= 6;
	private static byte		READLOCAL		= 7;
	private static byte		SSHLOGFILE		= 8;
	private static byte		SSHPRIVKEY		= 9;
	private static byte		STATS			= 10;
	private static byte		ROWMODE			= 11;
	private static byte		SANITIZE		= 12;
	private static byte		CSV				= 13;
	private static byte		HTML			= 14;
	private static byte		PROXY_TYPE		= 15;
	private static byte		PROXY_HOST		= 16;
	private static byte		PROXY_PORT		= 17;
	private static byte		PROXY_USER		= 18;
	private static byte		PROXY_PASSWORD	= 19;
	private static byte		SELECTED_MS		= 20;
	private static byte		CSV_SEPARATOR	= 21;
	private static byte		TIMEOUT			= 22;
	private static byte		CSVDIR			= 23;
	private static byte		HTMLDIR			= 24;
	private static byte		NOVIOS			= 25;
	private static byte		NAME			= 26;
	private static byte		NUM_CMD			= 27;
	
	
	
	
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
	
	// Management types
	
	private static final byte		M_HMC			= 0;
	private static final byte		M_SDMC			= 1;
	private static final byte		M_FSM			= 2;
	private static final byte		M_IVM			= 3;
	private static final byte		M_UNKNOWN		= 4;
	
	
	// Time period for perf data (days)
	private static final int		TIME_PERIOD		= 365; 
	
	// Managed system data
	
	private static final byte		CONFIGURABLE_POOL_PROC_UNITS		= 0;
	private static final byte		CURR_AVAIL_POOL_PROC_UNITS			= 1;
	private static final byte		BORROWED_POOL_PROC_UNITS			= 2;
	private static final byte		CONFIGURABLE_SYS_MEM				= 3;
	private static final byte		AVAIL_SYS_MEM						= 4;
	private static final byte		USED_POOL							= 5;
	private static final byte		NUM_MS								= 6;
	
	private float managedSystemData[][][] = null;	// [managedSystem][day][item]  day=0 means first day!!!
	private boolean goodSystemData[][] = null;		// [managedSystem][day]
	
	
	// LPAR data
	
	private static final byte		CURR_PROC_UNITS			= 0;
	private static final byte		CURR_PROCS				= 1;
	private static final byte		USED_CPU				= 2;
	private static final byte		NUM_LPAR				= 3;
	
	private String lparNames[] = null;
	private String lparID[] = null;
	private float lparData[][][] = null;			// [lpar][day][item]
	private boolean goodLparData[][] = null;		// [lpar][day]
	
	private static final byte		LPAR_TIME_CYCLES		= 0;
	private static final byte		LPAR_CYCLES				= 1;
	private static final byte		NUM_LPAR_CYCLES			= 2;
	
	private BigDecimal lpar_cycles[][][]=null;		// [lpar][day][item]
	
	
	// ManagedSystem historical config data
	private DataManager[] msCoreConfig = null;
	private DataManager[] msCoreAvail = null;
	private DataManager[] msCoreUsed = null;
	private DataManager[] msMemConfig = null;
	private DataManager[] msMemAvail = null;
	
	// LPAR historical config data
	private DataManager[] lparEnt = null;
	private DataManager[] lparVP = null;
	private DataManager[] lparPC = null;
	private NewLparStatus[]  lparStatus = null;
	
	// ProcPool historical config fata
	private DataManager[] procPoolConfig = null;
	private DataManager[] procPoolUsed = null;
	private String[] procPoolName = null;
	
	
	
	// Performance data summary
	
	private static final byte		WEEK1		= 0;
	private static final byte		WEEK2		= 1;
	private static final byte		WEEK3		= 2;
	private static final byte		WEEK4		= 3;
	private static final byte		MONTH1		= 4;
	private static final byte		MONTH2		= 5;
	private static final byte		MONTH3		= 6;
	private static final byte		MONTH4		= 7;
	private static final byte		MONTH5		= 8;
	private static final byte		MONTH6		= 9;
	private static final byte		MONTH7		= 10;
	private static final byte		MONTH8		= 11;	
	private static final byte		MONTH9		= 12;
	private static final byte		MONTH10		= 13;
	private static final byte		MONTH11		= 14;
	private static final byte		MONTH12		= 15;
	private static final byte		NUM_SUMMARY	= 16;
	
	private float managedSystemDataSummary[][][] = null;	// [managedSystem][sum][item]
	private float lparDataSummary[][][] = null;	// [lpar][sum][item]
	
	
	// Data type: dayly vs hourlu
	private static final byte		HOURLY		= 0;
	private static final byte		DAILY		= 1;
	
	
	

	private static final byte	UTILIZED_POOL_CYCLES 		= 0;
	private static final byte	TOTAL_POOL_CYCLES 			= 1;
	private static final byte	POOL_TIME_CYCLES			= 2;
	private static final byte	NUM_POOLDATA				= 3;
	
	private int Y_LEVEL = 70;
	private int R_LEVEL = 90;
	private double COLOR_LEVEL = 0.25;	// % of colored cells to set color in label
	
	private int XSIZE = 800;
	private int YSIZE = 400;
	
	
	private DiskData diskData = new DiskData();
	
	
	
	public Loader(String user, String password,  
					String startPerf, String endPerf, String samplePerf, boolean localOnly,
					String sshLogFile, String sshPrivKey, boolean produceStatistics, boolean rowMode, 
					String proxyType, String proxyHost, String proxyPort, String proxyUser, String proxyPassword,
					String selectedMS, String csvSeparatorString, int timeout, boolean novios) {
		//this.hmc = hmc;
		this.user = user;
		this.password = password;
		//this.baseDir = baseDir + File.separatorChar;
		this.startPerf = startPerf;
		this.endPerf = endPerf;
		this.samplePerf = samplePerf;
		this.onlyReadFile = localOnly;
		this.sshLogfile = sshLogFile;
		this.sshPrivKey = sshPrivKey;
		this.produceStatistics = produceStatistics;
		this.rowMode = rowMode;
		this.proxyType = proxyType;
		this.proxyHost = proxyHost;
		this.proxyPort = proxyPort;
		this.proxyUser = proxyUser;
		this.proxyPassword = proxyPassword;
		if (selectedMS==null)
			this.selectedMS=null;
		else
			this.selectedMS = selectedMS.replace(',','|');
		if (csvSeparatorString!=null)
			this.csvSeparator = csvSeparatorString.charAt(0);
		this.timeout = timeout;
		this.novios = novios;
		
		
		/*
		try {
			workbook = Workbook.createWorkbook(new File(baseDir+File.separatorChar+hmc+excel));
		} catch (IOException ioe) {
			System.out.println("Error in creating Excel file "+hmc+excel);
			System.exit(1);
		}
		*/
	}
	
	
	public boolean connect(String hmc) {
		//sshm = new SSHManager();
		sshm = new SSHManager2(sshLogfile);
		if (timeout!=0)
			sshm.setTimeout(timeout);
		
		if (sshPrivKey != null)
			return sshm.connectKey(hmc, user, sshPrivKey);
		
		if (proxyType!=null) {
			if (proxyType.equalsIgnoreCase("HTTP"))
				sshm.setProxyHTTP(proxyHost, Integer.parseInt(proxyPort), proxyUser, proxyPassword);
			if (proxyType.equalsIgnoreCase("SOCKS4"))
				sshm.setProxySOCKS4(proxyHost, Integer.parseInt(proxyPort), proxyUser, proxyPassword);
			if (proxyType.equalsIgnoreCase("SOCKS5"))
				sshm.setProxySOCKS5(proxyHost, Integer.parseInt(proxyPort), proxyUser, proxyPassword);
		}
		
		return sshm.connect(hmc, user, password);
		//return sshm.connectKey(hmc, user, "c:\\federico\\fed_id_rsa");
	}
	
	public void disconnect() {
		sshm.disconnect();
	}
	
	
	private static void showUsage() {
		showUsage(null);
	}
	
	
	private static void showUsage(String s) {
		System.out.println("HMC Scanner Version " + version);
		System.out.println("Missing or wrong arguments. Syntax is:");
		System.out.println("hmcScanner.Loader <HMC name/IP> <user> [-p <password>] [-dir <local dir>] [-perf <start> <end> <d|h>] [-readlocal] [-key [file] ] [-stats] [-sanitize] [-csv] [-csvchar <char>] [-html] [-log <file>] [-proxy <http|socks4|socks5> <host> <port> [user] [password] ] [-m <managed system list>] [-timeout <secconds>] [-novios] [-name <excelFIle>]");
		System.out.println("	-dir <local dir> is the directory where data will be stored. Default is current directory.");
		System.out.println("	-perf <start> and <end> is data collection retrieval interval. Syntax is: YYYYMMDD");
		System.out.println("	       <d|h> d=daily data samples; h=hourly data samples");
		System.out.println("	-readlocal will force reading of existing local data without contacting HMC");
		System.out.println("	-key will use OpenSSH private key (default $HOME/.ssh/id_rsa)");
		System.out.println("	-stats will produce system statistics. It does NOT require -perf!");
		//System.out.println("    -rowmode will show data by rows only (please select, it will be soon default!)");
		System.out.println("    -sanitize will remove sensitive data.");
		System.out.println("    -csv creates one .csv file for each sheet in a separate directory.");
		System.out.println("    -csvchar defines che separator char is CSV (default is comma).");
		System.out.println("    -html creates HTML report in a separate directory.");
		System.out.println("    -log <file> logs SSH activity into file");
		System.out.println("    -proxy <http|socks4|socks5> <host> <port> [user] [password]");
		System.out.println("           select proxy type, host, port to use for SSH communication");
		System.out.println("    -m <managed system list> restrics the scan to the provided comma separated managed system list");
		System.out.println("    -timeout <seconds> Timeout SSH connection after idle seconds. Use with care! Default is no timeout.");
		System.out.println("    -csvdir <directory> directory where csv files will be written. Default is <hmc>_<date>_<time>_CSV");
		System.out.println("    -htmldir <directory> directory where html files will be written. Default is <hmc>_<date>_<time>_HTML");
		System.out.println("    -novios do not collect data from VIOS to speed up pure hw config scanning.");
		System.out.println("    -name <excelFile> file name to be used for Excel output");
		System.out.println("");
		
		if (s!=null)
			System.out.println("Wrong argument is: "+s);
		System.exit(1);
	}
	
	
	private static String getPassword() {
		Console console = System.console();
	    if (console == null) {
	        System.out.println("Couldn't get Console instance");
	        System.exit(0);
	    }

	    console.printf("Testing password%n");
	    
	    char passwordArray[] = console.readPassword("Enter your password: ");
	    String password = new String(passwordArray);
	    
	    //console.printf("Password entered was: %s%n", password);
	    return password;
	}
	
	
	private static String[] parseParams(String[] args) {
		
		int i;
		
		// Minimum is HMC, user 
		if (args.length < 2)
			showUsage();
		for (i=0; i<2; i++)
			if (args[i].startsWith("-"))
				showUsage();
		
		String result[] = new String[NUM_CMD];
		for (i=0; i<NUM_CMD; i++)
			result[i]=null;
		
		result[HMC] 		= args[0];
		result[USER] 		= args[1];
		
		i=2;
		while (i<args.length) {
			
			// Password
			if (args[i].equals("-p")) {
				if (i+1>=args.length || args[i+1].startsWith("-"))
					showUsage();
				result[PASSWORD]=args[i+1];
				i+=2;
				continue;
			}
			
			// Local directory
			if (args[i].equals("-dir")) {
				if (i+1>=args.length || args[i+1].startsWith("-"))
					showUsage();
				result[DIR]=args[i+1];
				i+=2;
				continue;
			}
			
			// Collect performance
			if (args[i].equals("-perf")) {
				if (i+3>=args.length || args[i+1].startsWith("-") || args[i+2].startsWith("-") || args[i+3].startsWith("-"))
					showUsage();
				if (args[i+1].length()!=8 || args[i+2].length()!=8 || args[i+3].length()!=1)
					showUsage();
				int start=0;
				int end=0;
				try {
					start = Integer.parseInt(args[i+1]);
					end = Integer.parseInt(args[i+2]);
				} catch (NumberFormatException nfe) {
					showUsage();
				}
				if (start>end)
					showUsage();
				
				int month, day, year;
				
				month = start/100%100;
				day = start%100;
				year = start/10000;
				if (
						( month==2 && year%4!=0 && day>28 ) ||
						( month==2 && year%4==0 && day>29 ) ||
						( (month==1 || month==3 || month==5 || month==7 || month==8 || month==10 || month==12) && day>31 ) ||
						( (month==4 || month==6 || month==9 || month==11) && day>30 )			
				)
					showUsage();
				month = end/100%100;
				day = end%100;
				year = end/10000;
				if (
						( month==2 && year%4!=0 && day>28 ) ||
						( month==2 && year%4==0 && day>29 ) ||
						( (month==1 || month==3 || month==5 || month==7 || month==8 || month==10 || month==12) && day>31 ) ||
						( (month==4 || month==6 || month==9 || month==11) && day>30 )			
				)
					showUsage();
				
				if (!args[i+3].equals("d") && !args[i+3].equals("h"))
					showUsage();
					
				result[STARTPERF] 	= args[i+1];
				result[ENDPERF] 	= args[i+2];
				result[SAMPLE]		= args[i+3];	
				
				i+=4;
				continue;			
			}
			
			// Read only from local lata
			if (args[i].equals("-readlocal")) {
				result[READLOCAL]="true";
				i+=1;
				continue;
			}	
			
			// SSH logging
			if (args[i].equals("-log")) {
				if (i+1>=args.length || args[i+1].startsWith("-"))
					showUsage();
				result[SSHLOGFILE]=args[i+1];
				i+=2;
				continue;
			}
			
			// Timeout setting
			if (args[i].equals("-timeout")) {
				if (i+1>=args.length || args[i+1].startsWith("-"))
					showUsage();
				result[TIMEOUT]=args[i+1];
				i+=2;
				continue;
			}
			
			// SSH private key
			if (args[i].equals("-key")) {
				if (i+1>=args.length || args[i+1].startsWith("-"))
					showUsage();
				result[SSHPRIVKEY]=args[i+1];
				i+=2;
				continue;
			}
			
			// Collect 365 day data
			if (args[i].equals("-stats")) {
				result[STATS]="true";
				i+=1;
				continue;
			}
			
			// Show data in rows only ---> legacy: always true
			if (args[i].equals("-rowmode")) {
				result[ROWMODE]="true";
				i+=1;
				System.out.println("NOTE: -rowmode is now implicit and deprecated, please remove it from parameters.");
				continue;
			}
			
			// Sanitize data
			if (args[i].equals("-sanitize")) {
				result[SANITIZE]="true";
				i+=1;
				continue;
			}
			
			// CSV data
			if (args[i].equals("-csv")) {
				result[CSV]="true";
				i+=1;
				continue;
			}
			
			// HTML data
			if (args[i].equals("-html")) {
				result[HTML]="true";
				i+=1;
				continue;
			}
			
			// PROXY
			if (args[i].equals("-proxy")) {
				result[PROXY_TYPE]=args[++i];
				result[PROXY_HOST]=args[++i];
				result[PROXY_PORT]=args[++i];
				
				i+=1;
				
				if (i<args.length && !args[i].startsWith("-")) {
					result[PROXY_USER]=args[i];
								
					i+=1;
				
					if (i<args.length && !args[i].startsWith("-"))
						result[PROXY_PASSWORD]=args[i];
					
					i+=1;
				}
				
				if (!result[PROXY_TYPE].equalsIgnoreCase("HTTP") &&
						!result[PROXY_TYPE].equalsIgnoreCase("SOCKS4") &&
						!result[PROXY_TYPE].equalsIgnoreCase("SOCKS5") )
					showUsage();
				
				continue;
			}
			
			// Selected managed system 
			if (args[i].equals("-m")) {
				result[SELECTED_MS]=args[++i];
				i+=1;
				continue;
			}
			
			// Selected managed system 
			if (args[i].equals("-csvchar")) {
				result[CSV_SEPARATOR]=args[++i];
				i+=1;
				continue;
			}
			
			// Selected CSV dir 
			if (args[i].equals("-csvdir")) {
				result[CSVDIR]=args[++i];
				i+=1;
				continue;
			}
			
			// Selected HTML dir 
			if (args[i].equals("-htmldir")) {
				result[HTMLDIR]=args[++i];
				i+=1;
				continue;
			}	
			
			// NO DISK data
			if (args[i].equals("-novios")) {
				result[NOVIOS]="true";
				i+=1;
				continue;
			}
			
			// Selected XLS filename 
			if (args[i].equals("-name")) {
				result[NAME]=args[++i];
				i+=1;
				continue;
			}
			
			
			
			// Unexpected token
			showUsage(args[i]);
			
		}
		
		// New default is rowmode
		result[ROWMODE]="true";
		
		if (result[DIR]==null)
			result[DIR] = System.getProperty("user.dir");
		File dir = new File(result[DIR]);
		if (!dir.isDirectory() && !dir.mkdir()) {
			System.out.println("Error: can not create directory "+result[DIR]);
			System.exit(1);
		}
		/*
		result[DIR] = result[DIR] + File.separatorChar + result[HMC];
		dir = new File(result[DIR]);
		if (!dir.isDirectory() && !dir.mkdir()) {
			System.out.println("Error: can not create directory "+result[DIR]);
			System.exit(1);
		}	
		*/
		
		if (result[PASSWORD]==null && result[SSHPRIVKEY]==null) {
			String home = System.getenv("HOME");
			if (home!=null) {
				File f = new File(home+File.separatorChar+".ssh"+File.separatorChar+"id_rsa");
				if (f.exists() && f.canRead())
					result[SSHPRIVKEY]=home+File.separatorChar+".ssh"+File.separatorChar+"id_rsa";
				else
					//showUsage();
					result[PASSWORD]=getPassword();
			} else
				//showUsage();
				result[PASSWORD]=getPassword();
		}
		
		return result;		
	}
	
	
	public static void main(String[] args) {
		String result[];
		boolean onlyLocal = false;
		boolean produceStatistics = false;
		boolean rowMode = false;
		int timeout = 0;
		boolean novios=false;
		
		result = parseParams(args);
		if (result[READLOCAL]!=null && result[READLOCAL].equals("true"))
			onlyLocal = true;
		if (result[STATS]!=null && result[STATS].equals("true"))
			produceStatistics = true;
		if (result[ROWMODE]!=null && result[ROWMODE].equals("true"))
			rowMode=true;
		if (result[TIMEOUT]!=null)
			timeout = Integer.parseInt(result[TIMEOUT]);
		if (result[NOVIOS]!=null && result[NOVIOS].equals("true"))
			novios=true;
		
		Loader l = new Loader(result[USER], result[PASSWORD], 
								result[STARTPERF], result[ENDPERF], result[SAMPLE], onlyLocal,
								result[SSHLOGFILE], result[SSHPRIVKEY], produceStatistics, rowMode, 
								result[PROXY_TYPE], result[PROXY_HOST], result[PROXY_PORT], result[PROXY_USER], result[PROXY_PASSWORD],
								result[SELECTED_MS], result[CSV_SEPARATOR],timeout,novios);
		
		
		
		/*
		if (args.length<4 || args.length>5) 
			showUsage();
			
		if (args.length==5) {
			if (args[4].equals("true"))
				onlyLocal = true;
			else if (args[4].equals("false"))
				onlyLocal = false;
			else 
				showUsage();
		}
				
		Loader l = new Loader(args[0],args[1],args[2],args[3]+File.separatorChar,onlyLocal);
		
		*/
		
		//Loader l = new Loader("9.71.196.28","hscroot","abc1234","c:\\hmcScanner\\",true); 
		//l.connect();
		//l.getSystemData();
		
		File dir = new File(result[DIR]+File.separatorChar+result[HMC]);
		if (!dir.isDirectory() && !dir.mkdir()) {
			System.out.println("Error: can not create directory "+dir.getAbsolutePath());
			System.exit(1);
		}
		System.out.println("hmcScanner version "+version);
		l.startSession(result[HMC], result[DIR]+File.separatorChar+result[HMC]+File.separatorChar);		
		l.parseData(result[HMC], result[DIR]+File.separatorChar+result[HMC]+File.separatorChar);
		l.endSession();
		
		
		// If HTML output is required, create the destination directory
		String dirName;
		if (result[HTMLDIR]==null)
			dirName = result[DIR] + File.separatorChar + 
			result[HMC] + "_" +
			l.getHMCdate() + 
			html;
		else
			dirName = result[HTMLDIR];
		if (result[HTML]!=null && result[HTML].equals("true")) {
			File f = new File(dirName);
			if (!f.isDirectory() && !f.mkdir()) {
				System.out.println("Error: can not create directory "+dirName);
				return;
			}
		} else
			dirName=null;
		
		// If CSV output is required, create the destination directory
		String csvName = result[DIR] + File.separatorChar + 
							result[HMC] + "_" +
							l.getHMCdate() + 
							csv;
		if (result[CSVDIR]==null)
			csvName = result[DIR] + File.separatorChar + 
			result[HMC] + "_" +
			l.getHMCdate() + 
			csv;
		else
			csvName = result[CSVDIR];
		if (result[CSV]!=null && result[CSV].equals("true")) {
			File f = new File(csvName);
			if (!f.isDirectory() && !f.mkdir()) {
				System.out.println("Error: can not create directory "+csvName);
				return;
			}
		} else
			csvName=null;
		
		String excelName; 
		if (result[NAME]!=null) {
			excelName = result[NAME];
		} else
			excelName=		result[DIR] + File.separatorChar +
								result[HMC] + "_" +
								l.getHMCdate() + 
								excel;
		
		
		//l.createExcel(result[DIR]+File.separatorChar+"hmc_scanner_result.xls");
		l.createExcel(	excelName, dirName, csvName);
		
		/*
		if (result[CSV]!=null && result[CSV].equals("true")) 
			l.createCSVfiles(result[DIR]+File.separatorChar+"CSV", 
							result[DIR] + File.separatorChar +
							result[HMC] + "_" +
							l.getHMCdate() + 
							excel,
							result[HMC]);
		*/
		
		if (result[SANITIZE]!=null && result[SANITIZE].equals("true")) {
			l.sanitizeData();
			
			// If CSV output is required, create the destination directory
			if (result[CSVDIR]== null)
				csvName = result[DIR] + File.separatorChar + 
								result[HMC] + "_" +
								l.getHMCdate() + 
								csv_sane;
			else
				csvName = result[CSVDIR]+"_SANE";
			if (result[CSV]!=null && result[CSV].equals("true")) {
				File f = new File(csvName);
				if (!f.isDirectory() && !f.mkdir()) {
					System.out.println("Error: can not create directory "+csvName);
					return;
				}
			} else
				csvName=null;
			
			// If HTML output is required, create the destination directory
			if (result[HTMLDIR]== null)
				dirName = result[DIR] + File.separatorChar + 
								result[HMC] + "_" +
								l.getHMCdate() + 
								html_sane;
			else
				csvName = result[HTMLDIR]+"_SANE";
			if (result[HTML]!=null && result[HTML].equals("true")) {
				File f = new File(dirName);
				if (!f.isDirectory() && !f.mkdir()) {
					System.out.println("Error: can not create directory "+dirName);
					return;
				}
			} else
				dirName=null;
			
			if (result[NAME]!=null) {
				excelName = result[NAME].replaceAll("\\.xls$", "_sane.xls");
			} else
				excelName=		result[DIR] + File.separatorChar +
									result[HMC] + "_" +
									l.getHMCdate() + 
									excelSane;
			
			l.createExcel(	excelName, dirName, csvName);
			
			/*
			if (result[CSV]!=null && result[CSV].equals("true")) 
				l.createCSVfiles(result[DIR]+File.separatorChar+"CSV-SANE", 
								result[DIR] + File.separatorChar +
								result[HMC] + "_" +
								l.getHMCdate() + 
								excelSane,
								result[HMC]);
			*/
			
		}
		
		/*
		if (produceStatistics && result[HTML]!=null && result[HTML].equals("true")) {
			//l.createLparDailyFiles(result[DIR] + File.separatorChar + "HTML");
			try {
				l.createHTML(result[DIR] + File.separatorChar + 
								result[HMC] + "_" +
								l.getHMCdate() + 
								html);
			} catch (HeadlessException e) {	
				System.out.println("Warining: A graphical environment is required for HTML generation.");
				System.out.println("HTML graphical data skipped");				
			}
		}
		*/
	}
	
	
	private void startSession(String hmc, String baseDir) {
		if (!onlyReadFile) {
			if (!connect(hmc)) {
				System.out.println("Connection to HMC failed");
				System.exit(1);
			}
			collect_data(hmc, baseDir);
		} else {
			System.out.println("Reading local files. HMC will not be contacted.");
			identifyManagerType(baseDir);
		}
	}
	private void endSession() {
		if (!onlyReadFile)
			disconnect();
	}
	
	
	private String[] mergeList(String a[], String b[]) {
		String result[];
		String s[];
		int i,j,k, comp;
		
		
		result = new String[a.length];
		
		for (i=0; i<a.length; i++)
			result[i]=a[i];
		
		i=j=0;
		while (i<result.length && j<b.length) {
			comp = result[i].compareTo(b[j]);
			
			if (comp==0) {
				i++;
				j++;
				continue;
			}

			
			if (comp<0) {
				// b[j] is AFTER result[i]
				i++;
				continue;				
			}
			
			// b[j] is just BEFORE result[i]
			s=result;
			result=new String[s.length+1];
			for (k=0; k<i; k++)
				result[k]=s[k];
			result[i]=b[j];
			for (k=i+1; k<result.length; k++)
				result[k]=s[k-1];
			j++;	
		}
		
		if (j<b.length) {
			// Add remaining b components to result
			s=result;
			result=new String[s.length+b.length-j];
			for (k=0; k<s.length; k++)
				result[k]=s[k];
			for (k=s.length; k<result.length; k++)
				result[k]=b[j++];
		}
		
		return result;
	}
	
	
	
	/*
	 * Add a new GenericData to entPool[]
	 */
	private void addEntPool(GenericData pool) {
		if (entPool == null) {
			entPool = new GenericData[1];
			entPool[0] = pool;
			return;
		}
			
		GenericData newPool[] = new GenericData[entPool.length+1];
		int i;
		for (i=0; i<entPool.length; i++)
			newPool[i] = entPool[i];
		newPool[i] = pool;
		entPool = newPool;		
	}
	
	
	private void parseEntPool(String baseDir) {
		BufferedReader br;
		String names[]=null;
		String line;
		int j;
		DataParser dp;
		GenericData pool;
		
		try {			
			br = new BufferedReader(new FileReader(baseDir + lscodpool),1024*1024);
			
			names=null;
			while ( (line = br.readLine()) != null ) {
				
				// Skip line if no data is returned
				if (line.startsWith("HSC") || line.startsWith("No results were found"))
					continue;
				
				dp = new DataParser(line);
				names = dp.getNames();
				pool = new GenericData();
				
				for (j=0; j<names.length; j++) {
					pool.add(names[j], dp.getStringValue(names[j]));					
				}
				
				addEntPool(pool);			
			}		
			
		} 
		catch (IOException ioe) {	
			System.out.println("Loader.parseEntPool: IOException");
			System.out.println(ioe);
		}	
	}
	
	
	private void parseEntPoolSystem(String baseDir) {
		if (entPool == null)
			return;
		
		for (int i=0; i<entPool.length; i++) 
			load_data(entPool[i], baseDir, lscodpool, ENTPOOLSYS);		
	}
	
	
	
	private void detect_enterprise_pool(String prefix, String baseDir) {
		
		System.out.print("Looking for Enterprise System Pools: ");
		sendCommand(sshm, prefix + "lscodpool --level pool", baseDir + lscodpool);
		parseEntPool(baseDir);
		
		if (entPool == null)
			System.out.println("none detected");
		else {
			System.out.println(entPool.length + " detected");
			System.out.println("Scanning Enterprise pools: ");
			
			int i;
			String name;
			
			for (i=0; i<entPool.length; i++) {
				name = entPool[i].getVarValues("name")[0];
				sendCommand(sshm, prefix + "lscodpool --level sys -p " + name, baseDir + name + "_" + lscodpool);
				System.out.print(".");
				sendCommand(sshm, prefix + "lscodpool -t hist -p " + name, baseDir + name + "_hist_" + lscodpool);
				System.out.print(".");
			}
			
			System.out.println(" DONE");
		}	
	}
	
	
	
	private void collect_data(String hmc, String baseDir) {
		int		i,j;
		String	name;
		String  filler;
		String	prefix = "";
		int		n;
		
		
		
		FileWriter fw = null;
		GregorianCalendar cal = null;
		GregorianCalendar hmcDate = null;
		
		// Get date of HMC: if may not be the same of ours!
		//sshm.sendCommand("date +\"%Y-%m-%d %H:%M:%S\"", baseDir + hmcdate);
		sendCommand(sshm,"date +\"%Y-%m-%d %H:%M:%S\"", baseDir + hmcdate);
		hmcDate = parseHMCDate(baseDir + hmcdate);
		
		
		try {
			fw = new FileWriter(baseDir + scannerInfo);
			fw.write("HMC="+hmc+",");
			fw.write("user="+user+",");
			
			cal = new GregorianCalendar();
			
			fw.write("date="+cal.get(Calendar.YEAR)+"-");
			n = cal.get(Calendar.MONTH)+1;
			if (n<10) fw.write("0");
			fw.write(n+"-");
			n = cal.get(Calendar.DAY_OF_MONTH);
			if (n<10) fw.write("0");
			fw.write(n+",");
			
			fw.write("time=");
			n = cal.get(Calendar.HOUR_OF_DAY);
			if (n<10) fw.write("0");
			fw.write(n+":");
			n = cal.get(Calendar.MINUTE);
			if (n<10) fw.write("0");
			fw.write(n+":");
			n = cal.get(Calendar.SECOND);
			if (n<10) fw.write("0");
			fw.write(n+",");
			
			fw.write("HMCdate="+hmcDate.get(Calendar.YEAR)+"-");
			n = hmcDate.get(Calendar.MONTH)+1;
			if (n<10) fw.write("0");
			fw.write(n+"-");
			n = hmcDate.get(Calendar.DAY_OF_MONTH);
			if (n<10) fw.write("0");
			fw.write(n+",");
			
			fw.write("HMCtime=");
			n = hmcDate.get(Calendar.HOUR_OF_DAY);
			if (n<10) fw.write("0");
			fw.write(n+":");
			n = hmcDate.get(Calendar.MINUTE);
			if (n<10) fw.write("0");
			fw.write(n+":");
			n = hmcDate.get(Calendar.SECOND);
			if (n<10) fw.write("0");
			fw.write(n+"");
			
			/*
			fw.write("date="+cal.get(Calendar.YEAR)+"-"+(cal.get(Calendar.MONTH)+1)+"-"+cal.get(Calendar.DAY_OF_MONTH)+",");
			fw.write("time="+cal.get(Calendar.HOUR_OF_DAY)+":"+cal.get(Calendar.MINUTE)+":"+cal.get(Calendar.SECOND)+",");
			fw.write("HMCdate="+hmcDate.get(Calendar.YEAR)+"-"+(hmcDate.get(Calendar.MONTH)+1)+"-"+hmcDate.get(Calendar.DAY_OF_MONTH)+",");
			fw.write("HMCtime="+hmcDate.get(Calendar.HOUR_OF_DAY)+":"+hmcDate.get(Calendar.MINUTE)+":"+hmcDate.get(Calendar.SECOND));
			*/
			
			fw.flush();
			fw.close();
		} catch (IOException e) {
			System.out.println("Error in writing output files...");
			return;
		}
		
			
		System.out.print("Detecting manager type: ");
		//sshm.sendCommand("lshmc -V", baseDir + lshmcV);	
		sendCommand(sshm,"lshmc -V", baseDir + lshmcV);	
		identifyManagerType(baseDir);
		if (managerType == M_UNKNOWN) {
			//sshm.sendCommand("lsconfig -V", baseDir + lshmcV);
			sendCommand(sshm,"lsconfig -V", baseDir + lshmcV);
			identifyManagerType(baseDir);
		}
		if (managerType == M_UNKNOWN) {
			//sshm.sendCommand("ioscli ioslevel", baseDir + ivmversion);
			sendCommand(sshm,"ioscli ioslevel", baseDir + ivmversion);
			identifyManagerType(baseDir);
		}
		
		
		switch (managerType) {
			case M_HMC:		System.out.println("HMC");
							//sshm.sendCommand("lshmc -n", baseDir + lshmcn);
							//sshm.sendCommand("lshmc -b", baseDir + lshmcb);
							//sshm.sendCommand("lshmc -v", baseDir + lshmcv);
							sendCommand(sshm,"lshmc -n", baseDir + lshmcn);
							sendCommand(sshm,"lshmc -b", baseDir + lshmcb);
							sendCommand(sshm,"lshmc -v", baseDir + lshmcv);
							break;
							
			case M_SDMC:	System.out.println("SDMC");
							prefix = "smcli ";
							break;
							
			case M_FSM:		System.out.println("FSM");
							prefix = "smcli ";
							break;
							
			case M_IVM:		System.out.println("IVM");
							System.out.println("Warning: IVM is experimental...");
							break;
							
			default:		System.out.println("unknown... Aborting.");
							System.exit(1);
		}
		
		
	
		// Identify pools and download of their configuration
		if (managerType == M_HMC)
			detect_enterprise_pool(prefix, baseDir);
		
		System.out.print("Detecting managed systems: ");
		
		// Get utilization data configuration
		if (selectedMS!=null)
			//sshm.sendCommand(prefix + "lslparutil -r config -m "+selectedMS, baseDir + utilDataConfig);
			//sendCommand(sshm, prefix + "lslparutil -r config -m "+selectedMS, baseDir + utilDataConfig);
			sendCommand(sshm, prefix + "lslparutil -r config | egrep \""+selectedMS+"\"", baseDir + utilDataConfig);
		else
			//sshm.sendCommand(prefix + "lslparutil -r config", baseDir + utilDataConfig);
			sendCommand(sshm, prefix + "lslparutil -r config", baseDir + utilDataConfig);
		
		// Get system names and feature configuration
		if (selectedMS!=null)
			//sshm.sendCommand(prefix + "lssyscfg -r sys -m "+selectedMS, baseDir + systemData);
			//sendCommand(sshm, prefix + "lssyscfg -r sys -m "+selectedMS, baseDir + systemData);
			sendCommand(sshm, prefix + "lssyscfg -r sys | egrep \""+selectedMS+"\"", baseDir + systemData);
		else
			//sshm.sendCommand(prefix + "lssyscfg -r sys", baseDir + systemData);
			sendCommand(sshm, prefix + "lssyscfg -r sys", baseDir + systemData);
		
		// Get DLPAR info. No prefix|
		// sshm.sendCommand("lspartition -dlpar", baseDir + lspartition);
		
		// Load system data into parent data structure
		loadSysConfigData(hmc, baseDir);
		
		if (managedSystem == null) {
			System.out.println(" none detected!");
			return;
		}
		
		System.out.println(managedSystem.length+" systems present.");
		System.out.println("Starting managed system configuration collection:");
		
		for (i=0; i<managedSystem.length; i++) {
			name = managedSystem[i].getVarValues("name")[0];
			
			filler="";
			for (j=name.length(); j<28; j++)
				filler+=" ";
			
			System.out.print("   Scanning "+name+": "+filler);
			
			/*
			sshm.sendCommand(prefix + "lshwres -r proc     --level sys     -m \'"+name+"\'", baseDir + name + "_" + procSysData); 		System.out.print(".");			
			sshm.sendCommand(prefix + "lshwres -r mem      --level sys     -m \'"+name+"\'", baseDir + name + "_" + memSysData); 		System.out.print(".");
			sshm.sendCommand(prefix + "lshwres -r io       --rsubtype slot -m \'"+name+"\'", baseDir + name + "_" + slotSysData); 		System.out.print(".");
			sshm.sendCommand(prefix + "lshwres -r proc     --level lpar    -m \'"+name+"\'", baseDir + name + "_" + procLparData); 		System.out.print(".");
			sshm.sendCommand(prefix + "lshwres -r procpool                 -m \'"+name+"\'", baseDir + name + "_" + procPoolData); 		System.out.print(".");
			sshm.sendCommand(prefix + "lshwres -r mem      --level lpar    -m \'"+name+"\'", baseDir + name + "_" + memLparData); 		System.out.print(".");
			sshm.sendCommand(prefix + "lshwres -r mempool                  -m \'"+name+"\'", baseDir + name + "_" + memPoolData); 		System.out.print(".");
			sshm.sendCommand(prefix + "lssyscfg -r lpar                    -m \'"+name+"\'", baseDir + name + "_" + lparConfigData); 	System.out.print(".");
			sshm.sendCommand(prefix + "lssyscfg -r prof                    -m \'"+name+"\'", baseDir + name + "_" + lparProfiles); 	    System.out.print(".");
			sshm.sendCommand(prefix + "lslic                               -m \'"+name+"\'", baseDir + name + "_" + lslicSyspower); 			System.out.print(".");
			sshm.sendCommand(prefix + "lshwres -r virtualio --rsubtype vswitch          -m \'"+name+"\'", baseDir + name + "_" + vswitchData); 	System.out.print(".");
			sshm.sendCommand(prefix + "lshwres -r virtualio --rsubtype eth --level lpar -m \'"+name+"\'", baseDir + name + "_" + vethData); 	System.out.print(".");
			sshm.sendCommand(prefix + "lshwres -r virtualio --rsubtype scsi --level lpar -m \'"+name+"\'", baseDir + name + "_" + vscsiData); 	System.out.print(".");
			sshm.sendCommand(prefix + "lshwres -r virtualio --rsubtype fc --level lpar -m \'"+name+"\'", baseDir + name + "_" + vfcData); 		System.out.print(".");
			sshm.sendCommand(prefix + "lssyscfg -r lpar -m \'"+name+"\' -F name,lpar_env | grep vioserver", baseDir + name + "_" + viosList); System.out.print(".");
			sshm.sendCommand(prefix + "lshwres -m \'"+name+"\' -r io --rsubtype slotchildren", baseDir + name + "_" + slotchildren); System.out.print(".");
			
			sshm.sendCommand(prefix + "lscod -t bill -r proc -m \'"+name+"\'", baseDir + name + "_" + lscod_bill_proc); System.out.print(".");
			sshm.sendCommand(prefix + "lscod -t bill -r mem -m \'"+name+"\'", baseDir + name + "_" + lscod_bill_mem); System.out.print(".");
			sshm.sendCommand(prefix + "lscod -t cap -r proc -c onoff -m \'"+name+"\'", baseDir + name + "_" + lscod_cap_proc_onoff); System.out.print(".");
			sshm.sendCommand(prefix + "lscod -t cap -r mem -c onoff -m \'"+name+"\'", baseDir + name + "_" + lscod_cap_mem_onoff); System.out.print(".");
			sshm.sendCommand(prefix + "lscod -t hist -m \'"+name+"\'", baseDir + name + "_" + lscod_hist); System.out.print(".");
			*/
			
			sendCommand(sshm, prefix + "lshwres -r proc     --level sys     -m \'"+name+"\'", baseDir + name + "_" + procSysData); 		System.out.print(".");			
			sendCommand(sshm, prefix + "lshwres -r mem      --level sys     -m \'"+name+"\'", baseDir + name + "_" + memSysData); 		System.out.print(".");
			sendCommand(sshm, prefix + "lshwres -r io       --rsubtype slot -m \'"+name+"\'", baseDir + name + "_" + slotSysData); 		System.out.print(".");
			sendCommand(sshm, prefix + "lshwres -r proc     --level lpar    -m \'"+name+"\'", baseDir + name + "_" + procLparData); 		System.out.print(".");
			sendCommand(sshm, prefix + "lshwres -r procpool                 -m \'"+name+"\'", baseDir + name + "_" + procPoolData); 		System.out.print(".");
			sendCommand(sshm, prefix + "lshwres -r mem      --level lpar    -m \'"+name+"\'", baseDir + name + "_" + memLparData); 		System.out.print(".");
			sendCommand(sshm, prefix + "lshwres -r mempool                  -m \'"+name+"\'", baseDir + name + "_" + memPoolData); 		System.out.print(".");
			sendCommand(sshm, prefix + "lssyscfg -r lpar                    -m \'"+name+"\'", baseDir + name + "_" + lparConfigData); 	System.out.print(".");
			sendCommand(sshm, prefix + "lssyscfg -r prof                    -m \'"+name+"\'", baseDir + name + "_" + lparProfiles); 	    System.out.print(".");
			sendCommand(sshm, prefix + "lslic                               -m \'"+name+"\'", baseDir + name + "_" + lslicSyspower); 			System.out.print(".");
			sendCommand(sshm, prefix + "lshwres -r virtualio --rsubtype vswitch          -m \'"+name+"\'", baseDir + name + "_" + vswitchData); 	System.out.print(".");
			sendCommand(sshm, prefix + "lshwres -r virtualio --rsubtype eth --level lpar -m \'"+name+"\'", baseDir + name + "_" + vethData); 	System.out.print(".");
			sendCommand(sshm, prefix + "lshwres -r virtualio --rsubtype scsi --level lpar -m \'"+name+"\'", baseDir + name + "_" + vscsiData); 	System.out.print(".");
			sendCommand(sshm, prefix + "lshwres -r virtualio --rsubtype fc --level lpar -m \'"+name+"\'", baseDir + name + "_" + vfcData); 		System.out.print(".");
			sendCommand(sshm, prefix + "lssyscfg -r lpar -m \'"+name+"\' -F name,lpar_env | grep vioserver", baseDir + name + "_" + viosList); System.out.print(".");
			sendCommand(sshm, prefix + "lshwres -m \'"+name+"\' -r io --rsubtype slotchildren", baseDir + name + "_" + slotchildren); System.out.print(".");
			
			sendCommand(sshm, prefix + "lscod -t bill -r proc -m \'"+name+"\'", baseDir + name + "_" + lscod_bill_proc); System.out.print(".");
			sendCommand(sshm, prefix + "lscod -t bill -r mem -m \'"+name+"\'", baseDir + name + "_" + lscod_bill_mem); System.out.print(".");
			sendCommand(sshm, prefix + "lscod -t cap -r proc -c onoff -m \'"+name+"\'", baseDir + name + "_" + lscod_cap_proc_onoff); System.out.print(".");
			sendCommand(sshm, prefix + "lscod -t cap -r mem -c onoff -m \'"+name+"\'", baseDir + name + "_" + lscod_cap_mem_onoff); System.out.print(".");
			sendCommand(sshm, prefix + "lscod -t hist -m \'"+name+"\'", baseDir + name + "_" + lscod_hist); System.out.print(".");
			
			get_vios_lsmap_config(prefix, name, baseDir);	System.out.print(".");
			
			System.out.println(" DONE");
		}
		
		
		
		System.out.println("Collection successfully finished. Data is in "+baseDir);
		
		
		if (startPerf != null) {
			System.out.println("Starting performance data collection from all managed system:");
			
			for (i=0; i<managedSystem.length; i++) {
				name = managedSystem[i].getVarValues("name")[0];
				
				filler="";
				for (j=name.length(); j<28; j++)
					filler+=" ";
				
				System.out.print("   Loading "+name+": "+filler);
				
				/*
				sshm.sendCommand(
						prefix + 
						"lslparutil -r all -m \'" + name + "\'" + 
						" --startyear "  + startPerf.substring(0,4) +
						" --startmonth " + startPerf.substring(4,6) +
						" --startday "   + startPerf.substring(6)   +
						" --endyear "    + endPerf.substring(0,4)   +
						" --endmonth "   + endPerf.substring(4,6)   +
						" --endday "     + endPerf.substring(6)     +
						" -s " + samplePerf, 
						baseDir + name + "_" + lslparutilData,
						true);
				*/
				
				sendCommand( sshm, 
						prefix + 
						"lslparutil -r all -m \'" + name + "\'" + 
						" --startyear "  + startPerf.substring(0,4) +
						" --startmonth " + startPerf.substring(4,6) +
						" --startday "   + startPerf.substring(6)   +
						" --endyear "    + endPerf.substring(0,4)   +
						" --endmonth "   + endPerf.substring(4,6)   +
						" --endday "     + endPerf.substring(6)     +
						" -s " + samplePerf, 
						baseDir + name + "_" + lslparutilData,
						true);
				
				System.out.println(" DONE");
			}	
			
			System.out.println("Performance data collection ended. Data is in "+baseDir);
			System.out.println("Use pGraph.jar or other tools to read data.");
		}
		
		if (produceStatistics && managerType==M_IVM) {
			System.out.println("\n stats are buggy on IVM!");
			//produceStatistics=false;
			//return;
		}
		
		if (produceStatistics)
			updateStatistics(prefix, baseDir);	
	}
	
	
	private void updateStatistics(String prefix, String baseDir) {
		BufferedReader br = null;
		String line = null;
		int days = 0;
		int i,j;
		String name;
		String filler;
		
		
		/*
		try {
			br = new BufferedReader(new FileReader(baseDir + lastPerfDate),1024*1024);
			line=br.readLine();
			
			// 0123456789
			// 2012/12/31
			
			
			if (line==null || line.length()!=10) {
				// Empty or wrong line
				days = 365;
			} else {
				GregorianCalendar last = new GregorianCalendar(
												Integer.parseInt(line.substring(0,4)),
												Integer.parseInt(line.substring(5,7))-1,
												Integer.parseInt(line.substring(8)));
				//GregorianCalendar today = new GregorianCalendar();
				GregorianCalendar today = scannerDate;
				long diffMillis = today.getTimeInMillis()-last.getTimeInMillis();
				days = (int)(diffMillis/(24*60*60*1000));
			}		
		} catch (IOException ioe) {	
			// File does not exists
			days = 365;
		}
		*/
		
		// Always retrieve last 356 days!
		days = 365;
		
		System.out.println("Performance data collection: ");
		
		for (i=0; i<managedSystem.length; i++) {
			name = managedSystem[i].getVarValues("name")[0];
			
			filler="";
			for (j=name.length(); j<28; j++)
				filler+=" ";
			
			System.out.print("   Loading "+name+": "+filler);
			
			
			if (managerType == M_IVM) {
				
				GregorianCalendar day = new GregorianCalendar();
				day.add(Calendar.MONTH, -2);
				
				/*
				sshm.sendCommand(
						prefix + 
						"lslparutil -r sys " + 
						" --startyear " + day.get(Calendar.YEAR) +
						" --startmonth " + (day.get(Calendar.MONTH)+1) +
						" --startday " + day.get(Calendar.DAY_OF_MONTH) +
						" -n " + 2*30*24 +
						" --spread " + 
						" ; " +
						"lslparutil -r pool " + 
						" --startyear " + day.get(Calendar.YEAR) +
						" --startmonth " + (day.get(Calendar.MONTH)+1) +
						" --startday " + day.get(Calendar.DAY_OF_MONTH) +
						" -n " + 2*30*24 +
						" --spread " +
						" ; " +
						"lslparutil -r lpar " + 
						" --startyear " + day.get(Calendar.YEAR) +
						" --startmonth " + (day.get(Calendar.MONTH)+1) +
						" --startday " + day.get(Calendar.DAY_OF_MONTH) +
						" -n " + 2*30*24 +
						" --spread ", 
						baseDir + name + "_" + lslparutilHourlyStats,
						true, true); 
				*/
				
				sendCommand(sshm, 
						prefix + 
						"lslparutil -r sys " + 
						" --startyear " + day.get(Calendar.YEAR) +
						" --startmonth " + (day.get(Calendar.MONTH)+1) +
						" --startday " + day.get(Calendar.DAY_OF_MONTH) +
						" -n " + 2*30*24 +
						" --spread " + 
						" ; " +
						"lslparutil -r pool " + 
						" --startyear " + day.get(Calendar.YEAR) +
						" --startmonth " + (day.get(Calendar.MONTH)+1) +
						" --startday " + day.get(Calendar.DAY_OF_MONTH) +
						" -n " + 2*30*24 +
						" --spread " +
						" ; " +
						"lslparutil -r lpar " + 
						" --startyear " + day.get(Calendar.YEAR) +
						" --startmonth " + (day.get(Calendar.MONTH)+1) +
						" --startday " + day.get(Calendar.DAY_OF_MONTH) +
						" -n " + 2*30*24 +
						" --spread ", 
						baseDir + name + "_" + lslparutilHourlyStats,
						true, true); 
				
				day = new GregorianCalendar();
				day.add(Calendar.DAY_OF_YEAR, -days);
				
				/*
				sshm.sendCommand(
						prefix + 
						"lslparutil -r sys " + 
						" --startyear " + day.get(Calendar.YEAR) +
						" --startmonth " + (day.get(Calendar.MONTH)+1) +
						" --startday " + day.get(Calendar.DAY_OF_MONTH) +
						" -n " + days +
						" --spread " + 
						" ; " +
						"lslparutil -r pool " + 
						" --startyear " + day.get(Calendar.YEAR) +
						" --startmonth " + (day.get(Calendar.MONTH)+1) +
						" --startday " + day.get(Calendar.DAY_OF_MONTH) +
						" -n " + days +
						" --spread " +
						" ; " +
						"lslparutil -r lpar " + 
						" --startyear " + day.get(Calendar.YEAR) +
						" --startmonth " + (day.get(Calendar.MONTH)+1) +
						" --startday " + day.get(Calendar.DAY_OF_MONTH) +
						" -n " + days +
						" --spread ", 
						baseDir + name + "_" + lslparutilDailyStats,
						true, true);
				*/
				
				sendCommand(sshm, 
						prefix + 
						"lslparutil -r sys " + 
						" --startyear " + day.get(Calendar.YEAR) +
						" --startmonth " + (day.get(Calendar.MONTH)+1) +
						" --startday " + day.get(Calendar.DAY_OF_MONTH) +
						" -n " + days +
						" --spread " + 
						" ; " +
						"lslparutil -r pool " + 
						" --startyear " + day.get(Calendar.YEAR) +
						" --startmonth " + (day.get(Calendar.MONTH)+1) +
						" --startday " + day.get(Calendar.DAY_OF_MONTH) +
						" -n " + days +
						" --spread " +
						" ; " +
						"lslparutil -r lpar " + 
						" --startyear " + day.get(Calendar.YEAR) +
						" --startmonth " + (day.get(Calendar.MONTH)+1) +
						" --startday " + day.get(Calendar.DAY_OF_MONTH) +
						" -n " + days +
						" --spread ", 
						baseDir + name + "_" + lslparutilDailyStats,
						true, true);
				
				
			} else {
				/*
				sshm.sendCommand(
						prefix + 
						"lslparutil -r all -m \'" + name + "\'" + 
						" -d " + days +
						" -s h ", 
						baseDir + name + "_" + lslparutilHourlyStats,
						true, true);
				
				sshm.sendCommand(
						prefix + 
						"lslparutil -r all -m \'" + name + "\'" + 
						" -d " + days +
						" -s d ", 
						baseDir + name + "_" + lslparutilDailyStats,
						true, true);
				*/
				
				sendCommand(sshm, 
						prefix + 
						"lslparutil -r all -m \'" + name + "\'" + 
						" -d " + days +
						" -s h ", 
						baseDir + name + "_" + lslparutilHourlyStats,
						true, true);
				
				sendCommand(sshm, 
						prefix + 
						"lslparutil -r all -m \'" + name + "\'" + 
						" -d " + days +
						" -s d ", 
						baseDir + name + "_" + lslparutilDailyStats,
						true, true);
			}
			
			//System.out.println(" DONE");
			System.out.println("");
		}		
	}
	
	
	private void get_vios_lsmap_config(String prefix, String name, String baseDir) {
		BufferedReader br;
		String line;
		String vios;
		int comma;
		
		
		try {
			br = new BufferedReader(new FileReader(baseDir + name + "_" + viosList),1024*1024);
			
			while ( (line=br.readLine())!= null ) {
				
				comma = line.indexOf(',');
				if (comma<0)
					continue;
				
				vios = line.substring(0, comma);
				
				if (managerType == M_IVM) {
					/*
					sshm.sendCommand("ioscli ioslevel", baseDir + name + "_" + vios + "_" + ioslevel);
					sshm.sendCommand("ioscli lsmap -all -npiv -fmt :", baseDir + name + "_" + vios + "_" + npivData);	
					sshm.sendCommand("ioscli lsmap -all -fmt :", baseDir + name + "_" + vios + "_" + vscsiDiskData);
					sshm.sendCommand("ioscli chkdev -fmt : -field name identifier", baseDir + name + "_" + vios + "_" + diskuuid);
					sshm.sendCommand("ioscli lsdev | grep \"Shared Ethernet Adapter\" | while read i j ; do echo \"#$i\"; ioscli lsdev -dev $i -attr | grep True; done", baseDir + name + "_" + vios + "_" + seaCfg);
					sshm.sendCommand("ioscli lsdev | grep \"EtherChannel\" | while read i j ; do echo \"#$i\"; ioscli lsdev -dev $i -attr | grep True; done", baseDir + name + "_" + vios + "_" + etherChannel);
					sshm.sendCommand("ioscli lsdev | grep \"Shared Ethernet Adapter\" | while read i j ; do echo \"#$i\"; ioscli entstat -all $i ; done", baseDir + name + "_" + vios + "_" + entstatSEA);
					sshm.sendCommand("ioscli lsdev -vpd | grep -E \"^ *ent[0-9]+ +\"", baseDir + name + "_" + vios + "_" + slots);
					sshm.sendCommand("ioscli lsdev | grep -E \"^fcs[0-9]+ +\" | while read i j ; do echo \"#$i\"; ioscli fcstat -e $i ; done", baseDir + name + "_" + vios + "_" + fcstat);
					sshm.sendCommand("ioscli lsdev | grep -E \"^fcs[0-9]+ +\" | while read i j ; do echo \"#$i\"; ioscli lsdev -attr -dev $i ; done", baseDir + name + "_" + vios + "_" + fcattr);
					sshm.sendCommand("ioscli lspv -size -fmt :", baseDir + name + "_" + vios + "_" + lspv_size);
					sshm.sendCommand("ioscli lspv -free -fmt :", baseDir + name + "_" + vios + "_" + lspv_free);
					*/
					
					sendCommand(sshm, "ioscli ioslevel", baseDir + name + "_" + vios + "_" + ioslevel);
					sendCommand(sshm, "ioscli lsmap -all -npiv -fmt :", baseDir + name + "_" + vios + "_" + npivData);	
					sendCommand(sshm, "ioscli lsmap -all -fmt :", baseDir + name + "_" + vios + "_" + vscsiDiskData);
					sendCommand(sshm, "ioscli chkdev -fmt : -field name identifier", baseDir + name + "_" + vios + "_" + diskuuid);
					sendCommand(sshm, "ioscli lsdev | grep \"Shared Ethernet Adapter\" | while read i j ; do echo \"#$i\"; ioscli lsdev -dev $i -attr | grep True; done", baseDir + name + "_" + vios + "_" + seaCfg);
					sendCommand(sshm, "ioscli lsdev | grep \"EtherChannel\" | while read i j ; do echo \"#$i\"; ioscli lsdev -dev $i -attr | grep True; done", baseDir + name + "_" + vios + "_" + etherChannel);
					sendCommand(sshm, "ioscli lsdev | grep \"Shared Ethernet Adapter\" | while read i j ; do echo \"#$i\"; ioscli entstat -all $i ; done", baseDir + name + "_" + vios + "_" + entstatSEA);
					sendCommand(sshm, "ioscli lsdev -vpd | grep -E \"^ *ent[0-9]+ +\"", baseDir + name + "_" + vios + "_" + slots);
					sendCommand(sshm, "ioscli lsdev | grep -E \"^fcs[0-9]+ +\" | while read i j ; do echo \"#$i\"; ioscli fcstat -e $i ; done", baseDir + name + "_" + vios + "_" + fcstat);
					sendCommand(sshm, "ioscli lsdev | grep -E \"^fcs[0-9]+ +\" | while read i j ; do echo \"#$i\"; ioscli lsdev -attr -dev $i ; done", baseDir + name + "_" + vios + "_" + fcattr);
					sendCommand(sshm, "ioscli lspv -size -fmt :", baseDir + name + "_" + vios + "_" + lspv_size);
					sendCommand(sshm, "ioscli lspv -free -fmt :", baseDir + name + "_" + vios + "_" + lspv_free);
					sendCommand(sshm, "ioscli lsdev -dev proc0 -attr", baseDir + name + "_" + vios + "_" + proc0);
				} else if (!novios) {
					/*
					sshm.sendCommand(prefix + "viosvrcmd -m \'"+name+"\' -p \'"+vios+"\' -c \"ioslevel\"", baseDir + name + "_" + vios + "_" + ioslevel);					
					sshm.sendCommand(prefix + "viosvrcmd -m \'"+name+"\' -p \'"+vios+"\' -c \"lsmap -all -npiv -fmt :\"", baseDir + name + "_" + vios + "_" + npivData);	
					sshm.sendCommand(prefix + "viosvrcmd -m \'"+name+"\' -p \'"+vios+"\' -c \"lsmap -all -fmt :\"", baseDir + name + "_" + vios + "_" + vscsiDiskData);
					sshm.sendCommand(prefix + "viosvrcmd -m \'"+name+"\' -p \'"+vios+"\' -c \"chkdev -fmt : -field name identifier\"", baseDir + name + "_" + vios + "_" + diskuuid);
					sshm.sendCommand(prefix + "viosvrcmd -m \'"+name+"\' -p \'"+vios+"\' -c \"lsdev\" | grep \"Shared Ethernet Adapter\" | while read i j ; do echo \"#$i\"; " + 
											prefix + "viosvrcmd -m \'"+name+"\' -p \'"+vios+"\' -c \"lsdev -dev $i -attr\" | grep True ; done", baseDir + name + "_" + vios + "_" + seaCfg);
					sshm.sendCommand(prefix + "viosvrcmd -m \'"+name+"\' -p \'"+vios+"\' -c \"lsdev\" | grep \"EtherChannel\" | while read i j ; do echo \"#$i\"; " + 
							prefix + "viosvrcmd -m \'"+name+"\' -p \'"+vios+"\' -c \"lsdev -dev $i -attr\" | grep True ; done", baseDir + name + "_" + vios + "_" + etherChannel);
					sshm.sendCommand(prefix + "viosvrcmd -m \'"+name+"\' -p \'"+vios+"\' -c \"lsdev\" | grep \"Shared Ethernet Adapter\" | while read i j ; do echo \"#$i\"; " + 
							prefix + "viosvrcmd -m \'"+name+"\' -p \'"+vios+"\' -c \"entstat -all $i\" ; done", baseDir + name + "_" + vios + "_" + entstatSEA);
					sshm.sendCommand(prefix + "viosvrcmd -m \'"+name+"\' -p \'"+vios+"\' -c \"lsdev -vpd\" | grep -E \"^ *ent[0-9]+ +\"", baseDir + name + "_" + vios + "_" + slots);
					sshm.sendCommand(prefix + "viosvrcmd -m \'"+name+"\' -p \'"+vios+"\' -c \"lsdev\" | grep -E \"^fcs[0-9]+ +\" | while read i j ; do echo \"#$i\"; " + 
							prefix + "viosvrcmd -m \'"+name+"\' -p \'"+vios+"\' -c \"fcstat -e $i\" ; done", baseDir + name + "_" + vios + "_" + fcstat);
					sshm.sendCommand(prefix + "viosvrcmd -m \'"+name+"\' -p \'"+vios+"\' -c \"lsdev\" | grep -E \"^fcs[0-9]+ +\" | while read i j ; do echo \"#$i\"; " + 
							prefix + "viosvrcmd -m \'"+name+"\' -p \'"+vios+"\' -c \"lsdev -attr -dev $i\" ; done", baseDir + name + "_" + vios + "_" + fcattr);
					sshm.sendCommand(prefix + "viosvrcmd -m \'"+name+"\' -p \'"+vios+"\' -c \"lspv -size -fmt :\"", baseDir + name + "_" + vios + "_" + lspv_size);
					sshm.sendCommand(prefix + "viosvrcmd -m \'"+name+"\' -p \'"+vios+"\' -c \"lspv -free -fmt :\"", baseDir + name + "_" + vios + "_" + lspv_free);
					*/
					
					sendCommand(sshm, prefix + "viosvrcmd -m \'"+name+"\' -p \'"+vios+"\' -c \"ioslevel\"", baseDir + name + "_" + vios + "_" + ioslevel);					
					sendCommand(sshm, prefix + "viosvrcmd -m \'"+name+"\' -p \'"+vios+"\' -c \"lsmap -all -npiv -fmt :\"", baseDir + name + "_" + vios + "_" + npivData);	
					sendCommand(sshm, prefix + "viosvrcmd -m \'"+name+"\' -p \'"+vios+"\' -c \"lsmap -all -fmt :\"", baseDir + name + "_" + vios + "_" + vscsiDiskData);
					sendCommand(sshm, prefix + "viosvrcmd -m \'"+name+"\' -p \'"+vios+"\' -c \"chkdev -fmt : -field name identifier\"", baseDir + name + "_" + vios + "_" + diskuuid);
					sendCommand(sshm, prefix + "viosvrcmd -m \'"+name+"\' -p \'"+vios+"\' -c \"lsdev\" | grep \"Shared Ethernet Adapter\" | while read i j ; do echo \"#$i\"; " + 
											prefix + "viosvrcmd -m \'"+name+"\' -p \'"+vios+"\' -c \"lsdev -dev $i -attr\" | grep True ; done", baseDir + name + "_" + vios + "_" + seaCfg);
					sendCommand(sshm, prefix + "viosvrcmd -m \'"+name+"\' -p \'"+vios+"\' -c \"lsdev\" | grep \"EtherChannel\" | while read i j ; do echo \"#$i\"; " + 
								prefix + "viosvrcmd -m \'"+name+"\' -p \'"+vios+"\' -c \"lsdev -dev $i -attr\" | grep True ; done", baseDir + name + "_" + vios + "_" + etherChannel);
					sendCommand(sshm, prefix + "viosvrcmd -m \'"+name+"\' -p \'"+vios+"\' -c \"lsdev\" | grep \"Shared Ethernet Adapter\" | while read i j ; do echo \"#$i\"; " + 
							prefix + "viosvrcmd -m \'"+name+"\' -p \'"+vios+"\' -c \"entstat -all $i\" ; done", baseDir + name + "_" + vios + "_" + entstatSEA);
					sendCommand(sshm, prefix + "viosvrcmd -m \'"+name+"\' -p \'"+vios+"\' -c \"lsdev -vpd\" | grep -E \"^ *ent[0-9]+ +\"", baseDir + name + "_" + vios + "_" + slots);
					sendCommand(sshm, prefix + "viosvrcmd -m \'"+name+"\' -p \'"+vios+"\' -c \"lsdev\" | grep -E \"^fcs[0-9]+ +\" | while read i j ; do echo \"#$i\"; " + 
							prefix + "viosvrcmd -m \'"+name+"\' -p \'"+vios+"\' -c \"fcstat -e $i\" ; done", baseDir + name + "_" + vios + "_" + fcstat);
					sendCommand(sshm, prefix + "viosvrcmd -m \'"+name+"\' -p \'"+vios+"\' -c \"lsdev\" | grep -E \"^fcs[0-9]+ +\" | while read i j ; do echo \"#$i\"; " + 
							prefix + "viosvrcmd -m \'"+name+"\' -p \'"+vios+"\' -c \"lsdev -attr -dev $i\" ; done", baseDir + name + "_" + vios + "_" + fcattr);
					sendCommand(sshm, prefix + "viosvrcmd -m \'"+name+"\' -p \'"+vios+"\' -c \"lspv -size -fmt :\"", baseDir + name + "_" + vios + "_" + lspv_size);
					sendCommand(sshm, prefix + "viosvrcmd -m \'"+name+"\' -p \'"+vios+"\' -c \"lspv -free -fmt :\"", baseDir + name + "_" + vios + "_" + lspv_free);
					sendCommand(sshm, prefix + "viosvrcmd -m \'"+name+"\' -p \'"+vios+"\' -c \"lsdev -dev proc0 -attr\"", baseDir + name + "_" + vios + "_" + proc0);
				}
			}			
		} 
		catch (IOException ioe) {	
			System.out.println("Loader.get_vios_lsmap_config: IOException");
			System.out.println(ioe);
		}		
	}
	
	
	
	private void load_npiv_data(GenericData ms, String baseDir) {
		String sysName;
		BufferedReader br;
		String line;
		int comma;
		String vios;
		
		sysName = ms.getVarValues("name")[0];
		
		try {			
			br = new BufferedReader(new FileReader(baseDir + sysName + "_" + viosList),1024*1024);
			
			while ( (line=br.readLine())!= null ) {
				
				comma = line.indexOf(',');
				if (comma<0)
					continue;
				
				vios = line.substring(0, comma);
				
				load_vios_npiv_data(ms, vios, baseDir);
				load_vios_vscsi_data(ms, vios, baseDir);
				load_vios_ioslevel(ms, vios, baseDir);
				load_vios_hdisk_data(ms, vios, baseDir);
				
				load_sea_data(ms, vios, baseDir);
				
				load_fc_data(ms, vios, baseDir);
				
				load_freq_data(ms, vios, baseDir);
			}			
		} 
		catch (IOException ioe) {	
			System.out.println("Loader.load_npiv_data: IOException");
			System.out.println(ioe);
		}	
		
	}

	
	
	private void load_freq_data(GenericData ms, String viosName, String baseDir) {
		String sysName;
		BufferedReader br=null;
		String line;
		String values[]=null;
		String s[]=null;
		
		float freq;
		
		sysName = ms.getVarValues("name")[0];

		
		try {			
			br = new BufferedReader(new FileReader(baseDir + sysName + "_" + viosName + "_" + proc0),1024*1024);
			
			while ( (line=br.readLine())!= null ) {
				
				if (line.startsWith("HSC")) {
					if (verboseMode) {
						System.out.println("Warning: file " + baseDir + sysName + "_" + viosName + "_" + proc0 + " skipped due to invalid line");
						System.out.println("         Offending line = " + line);
					}
					return;
				}
				
				s = line.split("\\s+");
				if (s.length <2)
					continue;
				
				if (s[0].equals("frequency")) {			
					try {
						freq = Float.parseFloat(s[1]);
						freq = freq / 1000000000;
					} catch (NumberFormatException nfe) {
						System.out.println("Warning: could not evaluate CPU frequency from VIOS " + viosName + ": value=" + line);
						continue;
					}
					values = new String[1];
					values[0] = Float.toString(freq);
					ms.add("frequency", values);
					continue;
				}
				
				if (s[0].equals("type")) {			
					values = new String[1];
					values[0] = s[1];
					ms.add("cpu_type", values);
					continue;
				}
				
	
			}
			br.close();
		} 
		catch (IOException ioe) {	
			System.out.println("Loader.load_freq_data: IOException");
			System.out.println(ioe);
		}			
	}

	
	private void load_vios_ioslevel(GenericData ms, String viosName, String baseDir) {
		String sysName;
		BufferedReader br;
		String names[]=null;
		String line;
		GenericData lpars[];
		String values[]=null;
		
		sysName = ms.getVarValues("name")[0];
		
		try {			
			br = new BufferedReader(new FileReader(baseDir + sysName + "_" + viosName + "_" + ioslevel),1024*1024);
			
			while ( (line=br.readLine())!= null ) {
				
				if (line.startsWith("HSC")) {
					if (verboseMode) {
						System.out.println("Warning: file " + baseDir + sysName + "_" + viosName + "_" + ioslevel + " skipped due to invalid line");
						System.out.println("         Offending line = " + line);
					}
					return;
				}
				
				lpars = ms.getObjects(CONFIG_LPAR);
				for (int i=0; i<lpars.length; i++) {
					names = lpars[i].getVarValues("name");
					if (names[0].equals(viosName)) {
						values = new String[1];
						values[0]="VIOS "+line;
						lpars[i].add("os_version", values);
						return;
					}
				}			
			}			
		} 
		catch (IOException ioe) {	
			System.out.println("Loader.load_npiv_data: IOException");
			System.out.println(ioe);
		}		
		
	}

	
	
	/*
	 * Load NPIV data from file identified by viosName into a new set of objects in the objName category, and
	 * add such objects as children of ms, with viosName
	 */
	private void load_vios_npiv_data(GenericData ms, String viosName, String baseDir) {
		String sysName;
		BufferedReader br;
		String names[]=null;
		String line;
		GenericData gd;
		int begin, end;
		String viosSlot=null;
		int i,num;
		String newLine;
		
		sysName = ms.getVarValues("name")[0];
		
		try {			
			br = new BufferedReader(new FileReader(baseDir + sysName + "_" + viosName + "_" + npivData),1024*1024);
			
			while ( (line=br.readLine())!= null ) {
				
				// Return if no data is present
				if (line.startsWith("HSC") || line.startsWith(":") ) {
					if (verboseMode) {
						System.out.println("Warning: file " + baseDir + sysName + "_" + viosName + "_" + npivData + " skipped due to invalid line");
						System.out.println("         Offending line = " + line);
					}
					return;
				}
				
				gd = new GenericData();
				
				// Count number of ':'
				for (i=0, num=0; i<line.length(); i++)
					if (line.charAt(i)==':') num++;
				
				// Sometimes line is split. Try to recover
				if (num!=11) {
					System.out.println("Warning: npiv data line not in expected format -> " + line);
					newLine = br.readLine();
					if (newLine.startsWith("vfchost")) {
						System.out.println("NPIV data line skipped");
						line=newLine;
					} else {
						line = line + newLine;
						for (i=0, num=0; i<line.length(); i++)
							if (line.charAt(i)==':') num++;
						if (num==11) {
							System.out.println("NPIV line recevered!");
						} else {
							System.out.println("Warning: npiv line skipped with following one --> " + newLine);
							continue;
						}
					}
				}
				
				
				// Search virtual slot (parameter #2)
				begin = line.indexOf(':')+1;			// start of vfchost location 
				end = line.indexOf(':', begin);			// end of vfchost location
				begin = line.indexOf("-C", begin)+2;	// start of vfchost slot
				viosSlot=line.substring(begin, end);
				
				// Search physical adapter location (parameter #8)
				begin = line.indexOf(':',end+1);		 
				begin = line.indexOf(':',begin+1);		
				begin = line.indexOf(':',begin+1);		 
				begin = line.indexOf(':',begin+1);		
				begin = line.indexOf(':',begin+1)+1;		
				end = line.indexOf(':', begin);			// end of physical adapter location
				names = new String[1];
				names[0]=line.substring(begin, end);
				gd.add(viosName+"@"+viosSlot, names);
	
				// Add npiv mapping
				ms.addObject(VFCMAP, gd);			
			}			
		} 
		catch (IOException ioe) {	
			System.out.println("Loader.load_npiv_data: IOException");
			System.out.println(ioe);
		}			
		
	}
	
	
	/*
	 * Load sea data 
	 */
	private void load_sea_data(GenericData ms, String viosName, String baseDir) {
		String sysName;
		BufferedReader br;
		String names[]=null;
		String line;
		GenericData gd;
		int start, end;
		String varName, varValue;
		String tokens[];
	
		sysName = ms.getVarValues("name")[0];
		
		try {			
			br = new BufferedReader(new FileReader(baseDir + sysName + "_" + viosName + "_" + seaCfg),1024*1024);
			
			gd = null;
			
			while ( (line=br.readLine())!= null ) {
				
				// Return if no data is present
				if (line.startsWith("HSC") || line.startsWith(":") ) {
					if (verboseMode) {
						System.out.println("Warning: file " + baseDir + sysName + "_" + viosName + "_" + seaCfg + " skipped due to invalid line");
						System.out.println("         Offending line = " + line);
					}
					br.close();
					return;
				}
				
				// A line starting with # detects a new SEA
				if (line.startsWith("#")) {
					
					// Save previous SEA
					if (gd!=null)
						ms.addObject(SEA, gd);
					
					// Create a new SEA object
					gd = new GenericData();	
					
					names = new String[1];
					names[0]=viosName;
					gd.add("VIOS",names);
					
					names = new String[1];
					names[0]=line.substring(1);
					gd.add("SEA",names);	
					
					continue;
				}
				
				// Detect variable name
				end = line.indexOf(' ');
				varName = line.substring(0, end);
				
				start=end+1;
				while (line.charAt(start) == ' ')
					start++;
				
				end=line.indexOf(' ', start);
				varValue = line.substring(start, end);
				
				if (varValue.contains(",")) {
					// There are multiple values
					names = varValue.split(",");
					gd.add(varName, names);
				} else {
					names = new String[1];
					names[0]=varValue;
					if (varName.equals("ctl_chan") && !names[0].startsWith("ent"))
						names[0]="";
					gd.add(varName, names);
				}				
			}
			
			// Save previous SEA
			if (gd!=null)
				ms.addObject(SEA, gd);
		} 
		catch (IOException ioe) {	
			System.out.println("Loader.load_sea_data: IOException");
			System.out.println(ioe);
		}	
		
		
		try {			
			br = new BufferedReader(new FileReader(baseDir + sysName + "_" + viosName + "_" + slots),1024*1024);
			
			// Create a new ETH object
			gd = new GenericData();	
			
			names = new String[1];
			names[0]=viosName;
			gd.add("VIOS",names);
			
			while ( (line=br.readLine())!= null ) {
				
				// Return if no data is present
				if (line.startsWith("HSC") || line.startsWith(":") ) {
					if (verboseMode) {
						System.out.println("Warning: file " + baseDir + sysName + "_" + viosName + "_" + slots + " skipped due to invalid line");
						System.out.println("         Offending line = " + line);
					}
					br.close();
					return;
				}
				
				// Detect device name
				start = line.indexOf("ent");
				end = line.indexOf(' ', start);		
				varName = line.substring(start,end);
				
				// Detect slot name
				if (line.contains("Virtual")) {
					// It is a virtual eth
					start = line.indexOf("-C")+2;
					end = line.indexOf("-T1",start);
					varValue = line.substring(start, end);
				} else {
					// Physical or LEA
					start = line.indexOf("U");
					end = line.indexOf(' ',start);
					varValue = line.substring(start, end);
				}
				
				// Add data
				names = new String[1];
				names[0]=varValue;
				gd.add(varName,names);
			}
			
			// Save data
			ms.addObject(ETH, gd);
		} 
		catch (IOException ioe) {	
			System.out.println("Loader.load_sea_data: IOException");
			System.out.println(ioe);
		}	
		
		try {			
			br = new BufferedReader(new FileReader(baseDir + sysName + "_" + viosName + "_" + etherChannel),1024*1024);
			
			gd = null;
			
			while ( (line=br.readLine())!= null ) {
				
				// Return if no data is present
				if (line.startsWith("HSC") || line.startsWith(":") ) {
					if (verboseMode) {
						System.out.println("Warning: file " + baseDir + sysName + "_" + viosName + "_" + etherChannel + " skipped due to invalid line");
						System.out.println("         Offending line = " + line);
					}
					br.close();
					return;
				}
				
				// A line starting with # detects a new Etherchannel
				if (line.startsWith("#")) {
					
					// Save previous SEA
					if (gd!=null)
						ms.addObject(ETHERCHANNEL, gd);
					
					// Create a new SEA object
					gd = new GenericData();	
					
					names = new String[1];
					names[0]=viosName;
					gd.add("VIOS",names);
					
					names = new String[1];
					names[0]=line.substring(1);
					gd.add("ETHERCHANNEL",names);	
					
					continue;
				}
				
				// Detect variable name
				end = line.indexOf(' ');
				varName = line.substring(0, end);
				
				start=end+1;
				while (line.charAt(start) == ' ')
					start++;
				
				end=line.indexOf(' ', start);
				varValue = line.substring(start, end);
				
				if (varValue.contains(",")) {
					// There are multiple values
					names = varValue.split(",");
					gd.add(varName, names);
				} else {
					names = new String[1];
					names[0]=varValue;
					gd.add(varName, names);
				}				
			}
			
			// Save previous Etherchannel
			if (gd!=null)
				ms.addObject(ETHERCHANNEL, gd);
		} 
		catch (IOException ioe) {	
			System.out.println("Loader.load_sea_data: IOException");
			System.out.println(ioe);
		}
		
				
		
		try {			
			br = new BufferedReader(new FileReader(baseDir + sysName + "_" + viosName + "_" + entstatSEA),1024*1024);
			
			gd = null;
			
			GenericData vea=null;
			
			
			while ( (line=br.readLine())!= null ) {
				
				
				if (line.startsWith("#ent")) {
					// New SEA
					
					// Save previous VEA on previous SEA
					if (vea!=null) {
						gd.addObject(ETH, vea);
						vea = null;
					}
					
					// Save previous SEA
					if (gd!=null) 
						ms.addObject(ENTSTATSEA, gd);

					// Create a new SEA object
					gd = new GenericData();	
									
					names = new String[1];
					names[0]=viosName;
					gd.add("VIOS",names);
					
					names = new String[1];
					names[0]=line.substring(1);
					gd.add("SEA",names);	
									
					continue;				
				}
				
				if (line.startsWith("Virtual Adapter:") ||
						line.startsWith("Control Channel Adapter:")) {
					// Start of new VEA
					
					// Save previous VEA
					if (vea!=null)
						gd.addObject(ETH, vea);
					
					// Create new VEA
					vea = new GenericData();
					
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[tokens.length-1];
					vea.add("name",names);
					
					names = new String[1];
					if (tokens[0].equals("Control"))
						names[0]="Control";
					else
						names[0]="Data";
					vea.add("control_channel",names); 
					
					continue;
				}
				
				if (line.startsWith("Trunk")) {
					tokens = line.split("\\s+");
					if (tokens[tokens.length-1].equals("True"))  {
						line=br.readLine();
						if (line==null || !line.startsWith("  Priority:"))
							break;
						tokens = line.trim().split("\\s+");
						names = new String[2];
						names[0]=tokens[1];
						names[1]=tokens[3];
						vea.add("trunk",names);
					}
					continue;
				}
				
				if (line.startsWith("Hypervisor Send Failures:")) {
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[tokens.length-1];
					vea.add("hyp_send_failures",names);
					
					line=br.readLine();
					if (line==null || !line.startsWith("  Receiver Failures:"))
						break;
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[tokens.length-1];
					vea.add("hyp_send_receiver_fail",names);
					
					line=br.readLine();
					if (line==null || !line.startsWith("  Send Errors:"))
						break;
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[tokens.length-1];
					vea.add("hyp_send_sender_fail",names);
					
					line=br.readLine();
					if (line==null || !line.startsWith("Hypervisor Receive Failures:"))
						break;
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[tokens.length-1];
					vea.add("hyp_receive_failures",names);
					
					continue;
				}
				
				
				if (line.startsWith("Port VLAN ID:")) {
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[tokens.length-1];
					vea.add("PVID",names);
					
					line=br.readLine();
					if (line==null || !line.startsWith("VLAN Tag IDs:"))
						break;
					tokens = line.split("\\s+");
					
					names = new String[1];
					names[0]=tokens[3];
					for (int i=4; i<tokens.length; i++)
						names[0] = names[0] + ", " + tokens[i];
					
					// Next line is either empty or contains additional VLAN IDs
					line=br.readLine();
					if (line==null)
						break;
													
					if (line.length() > 5) {
						// additional vlans
						tokens = line.trim().split("\\s+");
						
						for (int i=0; i<tokens.length; i++)
							names[0] = names[0] + ", " + tokens[i];
					}
										
					vea.add("VLAN",names);
					
					continue;
				}
				
				
				if (line.startsWith("Switch ID:")) {
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[tokens.length-1];
					vea.add("switch",names);
					
					continue;
				}
				
				
				if (line.startsWith("    Buffer Size")) {
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[tokens.length-1];
					vea.add("bufsize",names);
					
					line=br.readLine();
					if (line==null || !line.startsWith("    Buffers"))
						break;
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[tokens.length-1];
					vea.add("numbufs",names);
					
					line=br.readLine();
					if (line==null || !line.startsWith("    History"))
						break;
					line=br.readLine();
					if (line==null || !line.startsWith("      No Buffers"))
						break;
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[tokens.length-1];
					vea.add("nobufs",names);
				}
				
				
				if (line.startsWith("    Min Buffers")) {
					tokens = line.trim().split("\\s+");
					names = new String[5];
					for (int i=0; i<5; i++)
						names[i]=tokens[2+i];
					vea.add("minBufs",names);
					
					line=br.readLine();
					if (line==null || !line.startsWith("    Max Buffers"))
						break;
					tokens = line.trim().split("\\s+");
					names = new String[5];
					for (int i=0; i<5; i++)
						names[i]=tokens[2+i];
					vea.add("maxBufs",names);
					
					line=br.readLine();
					if (line==null || !line.startsWith("    Allocated"))
						break;
					line=br.readLine();
					if (line==null || !line.startsWith("    Registered"))
						break;
					line=br.readLine();
					if (line==null || !line.startsWith("    History"))
						break;
					line=br.readLine();
					if (line==null || !line.startsWith("      Max Allocated"))
						break;
					
					tokens = line.trim().split("\\s+");
					names = new String[5];
					for (int i=0; i<5; i++)
						names[i]=tokens[2+i];
					vea.add("maxAlloc",names);
					
					continue;
				}
						
			}
			
			// Add VEA if any
			if (vea!=null)
				gd.addObject(ETH, vea);
			
			// Save SEA if any
			if (gd!=null)
				ms.addObject(ENTSTATSEA, gd);	
			
		} 
		catch (IOException ioe) {	
			System.out.println("Loader.load_sea_data: IOException");
			System.out.println(ioe);
		}
		
		
		/*
		try {			
			br = new BufferedReader(new FileReader(baseDir + sysName + "_" + viosName + "_" + entstatSEA),1024*1024);
			
			gd = null;
			
			boolean controlChannel;
			GenericData vea;
			
			line = skipToLine("#ent", br);
			if (line==null)
				return;
			
			do {
						
				// Save previous SEA
				if (gd!=null)
					ms.addObject(ENTSTATSEA, gd);
				
				// Create a new SEA object
				gd = new GenericData();	
								
				names = new String[1];
				names[0]=viosName;
				gd.add("VIOS",names);
				
				names = new String[1];
				names[0]=line.substring(1);
				gd.add("SEA",names);	
				
				controlChannel=false;
				while (  (line = skipToLine("#ent","Virtual Adapter:","Control Channel Adapter:",br)) != null) {
					
					if (line==null || line.startsWith("#ent"))
						break;
					
					if (line.startsWith("Control Channel Adapter"))
						controlChannel=true;
					
					vea = new GenericData();
					
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[tokens.length-1];
					vea.add("name",names);
					
					names = new String[1];
					if (controlChannel)
						names[0]="Control";
					else
						names[0]="Data";
					vea.add("control_channel",names);
					
					if (!controlChannel) {
						line = skipToLine("  Priority:",br);
						tokens = line.trim().split("\\s+");
						names = new String[2];
						names[0]=tokens[1];
						names[1]=tokens[3];
						vea.add("trunk",names);
					}
					
					line = skipToLine("Hypervisor Send Failures:",br);
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[tokens.length-1];
					vea.add("hyp_send_failures",names);
					
					line = skipToLine("  Receiver Failures:",br);
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[tokens.length-1];
					vea.add("hyp_send_receiver_fail",names);
					
					line = skipToLine("  Send Errors:",br);
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[tokens.length-1];
					vea.add("hyp_send_sender_fail",names);
					
					line = skipToLine("Hypervisor Receive Failures:",br);
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[tokens.length-1];
					vea.add("hyp_receive_failures",names);
					
					line = skipToLine("Port VLAN ID:",br);
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[tokens.length-1];
					vea.add("PVID",names);
					
					line = skipToLine("VLAN Tag IDs:",br);
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[3];
					for (int i=4; i<tokens.length; i++)
						names[0] = names[0] + ", " + tokens[i];
					vea.add("VLAN",names);
					
					line = skipToLine("Switch ID:",br);
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[tokens.length-1];
					vea.add("switch",names);
					
					line = skipToLine("    Buffer Size",br); if (line==null) break;
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[tokens.length-1];
					vea.add("bufsize",names);
					
					line = skipToLine("    Buffers",br); if (line==null) break;
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[tokens.length-1];
					vea.add("numbufs",names);
					
					line = skipToLine("      No Buffers",br); if (line==null) break;
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[tokens.length-1];
					vea.add("nobufs",names);
					
					line = skipToLine("    Min Buffers",br); if (line==null) break;
					tokens = line.trim().split("\\s+");
					names = new String[5];
					for (int i=0; i<5; i++)
						names[i]=tokens[2+i];
					vea.add("minBufs",names);
					
					line = skipToLine("    Max Buffers",br); if (line==null) break;
					tokens = line.trim().split("\\s+");
					names = new String[5];
					for (int i=0; i<5; i++)
						names[i]=tokens[2+i];
					vea.add("maxBufs",names);
					
					line = skipToLine("      Max Allocated",br); if (line==null) break;
					tokens = line.trim().split("\\s+");
					names = new String[5];
					for (int i=0; i<5; i++)
						names[i]=tokens[2+i];
					vea.add("maxAlloc",names);
					
					gd.addObject(ETH, vea);
				}
				
	
			} while ( line != null );
			
			// Save previous Etherchannel
			if (gd!=null)
				ms.addObject(ENTSTATSEA, gd);
		} 
		catch (IOException ioe) {	
			System.out.println("Loader.load_sea_data: IOException");
			System.out.println(ioe);
		}
		*/
		
	}
	
	
	
	/*
	 * Load FC data 
	 */
	private void load_fc_data(GenericData ms, String viosName, String baseDir) {
		String sysName;
		BufferedReader br;
		String names[]=null;
		String line;
		GenericData gd;
		int start, end;
		String tokens[];
		String varName, varValue;
		GenericData fc[];
		int	index;
		
		int valueCol=0, descrCol=0;

	
		sysName = ms.getVarValues("name")[0];
		
		try {			
			br = new BufferedReader(new FileReader(baseDir + sysName + "_" + viosName + "_" + fcattr),1024*1024);
			
			gd = null;
			
			while ( (line=br.readLine())!= null ) {
				
				// Return if no data is present
				if (line.startsWith("HSC") || line.startsWith(":") ) {
					if (verboseMode) {
						System.out.println("Warning: file " + baseDir + sysName + "_" + viosName + "_" + fcattr + " skipped due to invalid line");
						System.out.println("         Offending line = " + line);
					}
					br.close();
					return;
				}
				
				// A line starting with # detects a new FC
				if (line.startsWith("#")) {
					
					// Save previous FC
					if (gd!=null)
						ms.addObject(FC, gd); 
					
					// Create a new FC object
					gd = new GenericData();	
					
					names = new String[1];
					names[0]=viosName;
					gd.add("VIOS",names);
					
					names = new String[1];
					names[0]=line.substring(1);
					gd.add("FC",names);	
					
					continue;
				}
				
				// Skip lines
				if (line.length()==0)
					continue;
				
				// Skip header but detech label start
				if (line.length()==0 || line.startsWith("attribute")) {
					valueCol = line.indexOf("value");
					descrCol = line.indexOf("description");
					continue;
				}
				
				// Detect variable name
				/*
				end = line.indexOf(' ');
				varName = line.substring(0, end);
				
				start=end+1;
				while (line.charAt(start) == ' ')
					start++;
				
				end=line.indexOf(' ', start);
				varValue = line.substring(start, end);	
				*/
				varName = line.substring(0, valueCol).trim();
				varValue = line.substring(valueCol, descrCol).trim();
				
				names = new String[1];
				names[0]=varValue;
				gd.add(varName,names);	
			}
			
			// Save previous FC
			if (gd!=null)
				ms.addObject(FC, gd);
		} 
		catch (IOException ioe) {	
			System.out.println("Loader.load_fc_data: IOException");
			System.out.println(ioe);
		}	
		
				
		
		try {			
			br = new BufferedReader(new FileReader(baseDir + sysName + "_" + viosName + "_" + fcstat),1024*1024);
			
			gd = null;			
			
			while ( (line=br.readLine())!= null ) {
				
				
				if (line.startsWith("#fc")) {
					// New FC
					
					
					// Search existing FC object
					fc = ms.getObjects(FC);
					for (index=0; index<fc.length; index++)
						if (fc[index].getVarValues("VIOS")[0].equals(viosName) &&
								fc[index].getVarValues("FC")[0].equals(line.substring(1)) )
							break;
					if (index==fc.length) {
						// Unlikely to happen, but...
						// Create a new FC object
						gd = new GenericData();	
										
						names = new String[1];
						names[0]=viosName;
						gd.add("VIOS",names);
						
						names = new String[1];
						names[0]=line.substring(1);
						gd.add("FC",names);	
						ms.addObject(FC, gd); 				
					} else
						gd = fc[index];
									
					continue;				
				}
				
				if (line.startsWith("ZA:")) {
					// Firmware
					tokens = line.split("\\s+");
					names = new String[1];
					if (tokens.length <2) {
						// no firmware is provided
						names[0]="N/A";
						gd.add("firmware",names);
					} else {
						names[0]=tokens[1];
						gd.add("firmware",names);
					}
					continue;
				}
				
				if (line.startsWith("World Wide Node Name:")) {
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[tokens.length-1];
					gd.add("wwnn",names);
					continue;
				}
				
				if (line.startsWith("World Wide Port Name:")) {
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[tokens.length-1];
					gd.add("wwpn",names);
					continue;
				}
				
				if (line.startsWith("Port Speed (supported):")) {
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[tokens.length-2];
					gd.add("speed-supported",names);
					continue;
				}
				
				if (line.startsWith("Port Speed (running):")) {
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[tokens.length-2];
					gd.add("speed-running",names);
					continue;
				}
				
				if (line.startsWith("Port FC ID:")) {
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[tokens.length-1];
					gd.add("fcid",names);
					continue;
				}
				
				if (line.startsWith("Port Type:")) {
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[tokens.length-1];
					gd.add("port_type",names);
					continue;
				}
				
				if (line.startsWith("Seconds Since Last Reset:")) {
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[tokens.length-1];
					gd.add("reset",names);
					continue;
				}
				
				if (line.startsWith("Error Frames:")) {
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[tokens.length-1];
					gd.add("errorf",names);
					continue;
				}
				
				if (line.startsWith("Dumped Frames:")) {
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[tokens.length-1];
					gd.add("dumpedf",names);
					continue;
				}
				
				if (line.startsWith("Invalid CRC Count:")) {
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[tokens.length-1];
					gd.add("invalidcrc",names);
					continue;
				}
				
				if (line.startsWith("Invalid Tx Word Count:")) {
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[tokens.length-1];
					gd.add("invalidtx",names);
					continue;
				}
				
				if (line.startsWith("IP over FC Adapter Driver Information")) {
					line=br.readLine();
					if (line==null || !line.startsWith("  No DMA Resource Count:"))
						break;
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[tokens.length-1];
					gd.add("ip_nodma",names);
					
					line=br.readLine();
					if (line==null || !line.startsWith("  No Adapter Elements Count:"))
						break;
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[tokens.length-1];
					gd.add("ip_noadapter",names);
					continue;
				}
				
				if (line.startsWith("FC SCSI Adapter Driver Information")) {
					line=br.readLine();
					if (line==null || !line.startsWith("  No DMA Resource Count:"))
						break;
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[tokens.length-1];
					gd.add("scsi_nodma",names);
					
					line=br.readLine();
					if (line==null || !line.startsWith("  No Adapter Elements Count:"))
						break;
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[tokens.length-1];
					gd.add("scsi_noadapter",names);
					
					line=br.readLine();
					if (line==null || !line.startsWith("  No Command Resource Count:"))
						break;
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[tokens.length-1];
					gd.add("scsi_nocmd",names);
					
					continue;
				}
				
				if (line.startsWith("Adapter Effective max transfer value:")) {
					tokens = line.split("\\s+");
					names = new String[1];
					names[0]=tokens[tokens.length-1];
					gd.add("e_max_transfer",names);
					continue;
				}
						
			}
			
			// Save FCSTAT if any
			if (gd!=null)
				ms.addObject(FCSTAT, gd);	
			
		} 
		catch (IOException ioe) {	
			System.out.println("Loader.load_fc_data: IOException");
			System.out.println(ioe);
		}
		
	}
	
	
	private String skipToLine(String s, BufferedReader br) throws IOException {
		String line;
		
		while ( (line=br.readLine())!= null ) {
			// Return if no data is present
			if (line.startsWith("HSC") || line.startsWith(":") )
				return null;
			
			if (line.startsWith(s))
				return line;
		}
		return null;
	}
	
	private String skipToLine(String s1, String s2, String s3, BufferedReader br) throws IOException {
		String line;
		
		while ( (line=br.readLine())!= null ) {
			// Return if no data is present
			if (line.startsWith("HSC") || line.startsWith(":") )
				return null;
			
			if (line.startsWith(s1) || line.startsWith(s2) || line.startsWith(s3))
				return line;
		}
		return null;
	}
	
	
	
	/*
	 * Load hdisk data from chkdev
	 */
	private void load_vios_hdisk_data(GenericData ms, String viosName, String baseDir) {
		String sysName;
		BufferedReader br;
		String names[]=null;
		String line;
		GenericData gd;
		int end;
		String hdisk; 
		String uuid;
		
		sysName = ms.getVarValues("name")[0];
		
		try {			
			br = new BufferedReader(new FileReader(baseDir + sysName + "_" + viosName + "_" + diskuuid),1024*1024);
			
			gd = new GenericData();
			
			names = new String[1];
			names[0]=viosName;
			gd.add("VIOS",names);
			
			while ( (line=br.readLine())!= null ) {
				
				// Return if no data is present
				if (line.startsWith("HSC") || line.startsWith(":") ) {
					if (verboseMode) {
						System.out.println("Warning: file " + baseDir + sysName + "_" + viosName + "_" + diskuuid + " skipped due to invalid line");
						System.out.println("         Offending line = " + line);
					}
					br.close();
					return;
				}
						
				// Search virtual slot (parameter #2)
				end = line.indexOf(':');
				hdisk=line.substring(0, end);				
				uuid=line.substring(end+1);

				names = new String[1];
				names[0]=uuid;
				gd.add(hdisk, names);
				
				diskData.add_uuid(uuid, hdisk, sysName, viosName);
				
			}
			
			ms.addObject(HDISK, gd);
		} 
		catch (IOException ioe) {	
			System.out.println("Loader.load_npiv_data: IOException");
			System.out.println(ioe);
		}	
		
		

		try {			
			br = new BufferedReader(new FileReader(baseDir + sysName + "_" + viosName + "_" + lspv_size),1024*1024);
						
			String size;
			
			while ( (line=br.readLine())!= null ) {
				
				// Return if no data is present
				if (line.startsWith("HSC") || line.startsWith(":") ) {
					if (verboseMode) {
						System.out.println("Warning: file " + baseDir + sysName + "_" + viosName + "_" + lspv_size + " skipped due to invalid line");
						System.out.println("         Offending line = " + line);
					}
					br.close();
					return;
				}
						
				end = line.indexOf(':');
				hdisk=line.substring(0, end);	
				end = line.indexOf(':', end+1);			
				size=line.substring(end+1);
				
				diskData.addSize(sysName, viosName, hdisk, Integer.valueOf(size));				
			}
		} 
		catch (IOException ioe) {	
			System.out.println("Loader.load_npiv_data: IOException");
			System.out.println(ioe);
		}
		
		try {			
			br = new BufferedReader(new FileReader(baseDir + sysName + "_" + viosName + "_" + lspv_free),1024*1024);
									
			while ( (line=br.readLine())!= null ) {
				
				// Return if no data is present
				if (line.startsWith("HSC") || line.startsWith(":") ) {
					if (verboseMode) {
						System.out.println("Warning: file " + baseDir + sysName + "_" + viosName + "_" + lspv_free + " skipped due to invalid line");
						System.out.println("         Offending line = " + line);
					}
					br.close();
					return;
				}
							
						
				end = line.indexOf(':');
				hdisk=line.substring(0, end);	
				
				diskData.addFree(sysName, viosName, hdisk);				
			}
		} 
		catch (IOException ioe) {	
			System.out.println("Loader.load_npiv_data: IOException");
			System.out.println(ioe);
		}
		
	}
	
	
	/*
	 * Load VSCSI data from file identified by viosName into a new set of objects in the objName category, and
	 * add such objects as children of ms, with viosName
	 */
	private void load_vios_vscsi_data(GenericData ms, String viosName, String baseDir) {
		String sysName;
		BufferedReader br;
		String names[]=null;
		String line;
		GenericData gd;
		int begin, end;
		
		String data[]=null;
		int	index;
		
		
		String SVSA = null;
		String slot = null;
		String client = null;
		Vector<String> VTD = null;
		Vector<String> status = null;
		Vector<String> LUN = null;
		Vector<String> backingDevice = null;
		Vector<String> physloc = null;
		Vector<String> mirror = null;
		
		int numparams = 0;
		int possibleStatus = 0;
		String s;
		
		sysName = ms.getVarValues("name")[0];
		
		try {			
			br = new BufferedReader(new FileReader(baseDir + sysName + "_" + viosName + "_" + vscsiDiskData),1024*1024);
			
			while ( (line=br.readLine())!= null ) {
				
				//System.out.println(baseDir + sysName + "_" + viosName + "_" + vscsiDiskData + "\t" + line);
				
				// Return if no data is present
				if (line.startsWith("HSC") || line.startsWith(":") ) {
					if (verboseMode) {
						System.out.println("Warning: file " + baseDir + sysName + "_" + viosName + "_" + vscsiDiskData + " skipped due to invalid line");
						System.out.println("         Offending line = " + line);
					}
					br.close();
					return;
				}
				
				if (!line.contains(":"))
					continue;
				
				data = line.split(":");
				if (line.contains(":true") || line.contains(":false") || line.contains(":N/A"))
					numparams = 6; // 6 params, including mirror
				else if (line.contains(":Available") || line.contains(":Defined"))
						numparams = 5;
					else
						numparams = 4;
				
				gd = new GenericData();
				
				// Search SVSA (parameter #1)
				SVSA = data[0];
				
				// Search slot number (parameter #2)
				begin = data[1].indexOf("-C");
				slot=data[1].substring(begin+2);
				
				// Search client identifier
				client = data[2];				
				
				// Scan mapped devices				
				VTD = new Vector<String>();
				status = new Vector<String>();
				LUN = new Vector<String>();
				backingDevice = new Vector<String>();
				physloc = new Vector<String>();
				mirror = new Vector<String>();
				
				index=3;
				while (index<=data.length-numparams) {					
					//VTD
					VTD.add(data[index++]);
					//Status
					if (numparams!=4) {	status.add(data[index++]);	}	
					//LUN
					LUN.add(data[index++]);
					//backingDevice
					backingDevice.add(data[index++]);
					//physloc 
					physloc.add(data[index++]);
					//mirror
					if (numparams == 6) { mirror.add(data[index++]); }
				}
				
				
				
				/* OLD OLD OLD 
				
				numparams = 0;
				for (int i=0; i<line.length(); i++)
					if (line.charAt(i)==':') {
						numparams++;
						if (numparams == 4)
							possibleStatus = i+1;
					}
				
				if (line.contains(":true") || line.contains(":false") || line.contains(":N/A"))
					numparams = 6; // 6 params, including mirror
				else {
					s = line.substring(possibleStatus, line.indexOf(':', possibleStatus));
					if (s.equals("Available") || s.equals("Defined"))
						numparams = 5;
					else
						numparams = 4;
				}
				

				
				gd = new GenericData();
				
				// Search SVSA (parameter #1)
				end = line.indexOf(':');			
				SVSA = line.substring(0, end);
				begin = end + 1;
				
				// Search slot number (parameter #2)
				begin = line.indexOf("-C", begin);
				end = line.indexOf(":",begin);
				slot=line.substring(begin+2, end);
				begin = end + 1;
				
				// Search client identifier
				end = line.indexOf(":",begin);
				client = line.substring(begin, end);
				begin = end + 1;
				
				
				// Scan mapped devices				
				VTD = new Vector<String>();
				status = new Vector<String>();
				LUN = new Vector<String>();
				backingDevice = new Vector<String>();
				physloc = new Vector<String>();
				mirror = new Vector<String>();
				
				//while (!line.endsWith(": : ") && !line.endsWith("::") && end<line.length()) {
				while (end<line.length()) {
					
					// Check if empty line
					end = line.indexOf(":",begin);
					if (end == -1)
						break;		// truncated line
					if (end==begin)
						break;		// "::" is detected
					if (end==begin+1 && line.charAt(begin)==' ')
						break;		// ": :" is detected
					
					//VTD
					end = line.indexOf(":",begin);
					VTD.add(line.substring(begin, end));
					begin = end + 1;
					
					//Status
					if (numparams!=4) {
						end = line.indexOf(":",begin);
						status.add(line.substring(begin, end));
						begin = end + 1;
					}
					
					//LUN
					end = line.indexOf(":",begin);
					LUN.add(line.substring(begin, end));
					begin = end + 1;
					
					//backingDevice
					end = line.indexOf(":",begin);
					backingDevice.add(line.substring(begin, end));
					begin = end + 1;
					
					//physloc (it may be the last)
					end = line.indexOf(":",begin);
					if (end>0)
						physloc.add(line.substring(begin, end));
					else {
						physloc.add(line.substring(begin));
						end = line.length();
					}
					begin = end + 1;
					
					
					if (numparams == 6) {
						//mirror
						end = line.indexOf(":",begin);
						if (end>0)
							mirror.add(line.substring(begin, end));
						else {
							mirror.add(line.substring(begin));
							end = line.length();
						}
						begin = end + 1;
					}	
				}
				*/
				
				names = new String[1];
				names[0]=viosName;
				gd.add("VIOS", names);
				
				names = new String[1];
				names[0]=slot;
				gd.add("slot", names);
				
				names = new String[1];
				names[0]=SVSA;
				gd.add("SVSA", names);
					
				names = new String[1];
				names[0]=client;
				gd.add("client", names);
				
				names = new String[1];
				gd.add("VTD", 			VTD.toArray(names));
				names = new String[1];
				gd.add("Status", 		status.toArray(names));
				names = new String[1];
				gd.add("LUN", 			LUN.toArray(names));
				names = new String[1];
				gd.add("BackingDevice", backingDevice.toArray(names));
				names = new String[1];
				gd.add("physloc", 		physloc.toArray(names));	
				names = new String[1];
				gd.add("mirror", 		mirror.toArray(names));
	
				// Add VSCSI mapping
				ms.addObject(VSCSIMAP, gd);			
			}			
		} 
		catch (IOException ioe) {	
			System.out.println("Loader.load_vios_vscsi_data: IOException");
			System.out.println(ioe);
		}			
		
	}
	
	
	
	
	private GregorianCalendar parseHMCDate(String filename) {
		BufferedReader br;

		String line;
		GregorianCalendar result=null;
		int Y,M,D;
		int hh,mm,ss;
		int s,e;
		
	
		try {			
			br = new BufferedReader(new FileReader(filename),1024*1024);
			
			s=e=0;
			while ( (line=br.readLine())!= null ) {
				e = line.indexOf('-');
				Y=Integer.parseInt(line.substring(0, e));
				s=e+1; e=line.indexOf('-',s);
				M=Integer.parseInt(line.substring(s, e));
				s=e+1; e=line.indexOf(' ',s);
				D=Integer.parseInt(line.substring(s,e));
				s=e+1; e=line.indexOf(':',s);
				hh=Integer.parseInt(line.substring(s,e));
				s=e+1; e=line.indexOf(':',s);
				mm=Integer.parseInt(line.substring(s,e));
				s=e+1;
				ss=Integer.parseInt(line.substring(s));
				result = new GregorianCalendar(Y,M-1,D,hh,mm,ss);			
			}	
			br.close();
			return result;
		} 
		catch (Exception exc) {	
			System.out.println("Loader.parseHMCDate Exception");
			System.out.println(exc);
			return null;
		}			
		
	}
	
	
	
	
	
	/*
	 * Load file identified by fileName into a new set of objects in the objName category, and
	 * add such objects as children of ms
	 */
	private void load_data(GenericData ms, String baseDir, String fileName, byte objName) {
		String sysName;
		BufferedReader br=null;
		String names[]=null;
		String line;
		DataParser dp;
		GenericData gd;
		int i;
		
		sysName = ms.getVarValues("name")[0];
		
		try {			
			br = new BufferedReader(new FileReader(baseDir + sysName + "_" + fileName),1024*1024);
			
			while ( (line=br.readLine())!= null ) {
				
				// Skip line if no data is returned
				if (line.startsWith("HSC") || line.startsWith("The managed system") || line.startsWith("No results were found")) {
					if (verboseMode) {
						System.out.println("Warning: file " + baseDir + sysName + "_" + fileName + " skipped due to invalid line");
						System.out.println("         Offending line = " + line);
					}
					br.close();
					return;
				}
				
				dp = new DataParser(line);
				names = dp.getNames();
				
				// Skip line if no variable have been identified
				if (names==null || names.length==0 || names[0]==null)
					continue;
				
				gd = new GenericData();
				
				for (i=0; i<names.length; i++) {
					gd.add(names[i], dp.getStringValue(names[i]));					
				}	
				
				ms.addObject(objName, gd);			
			}	
		} 
		catch (IOException ioe) {	
			System.out.println("Loader.load_data: IOException");
			System.out.println(ioe);
		}	
	}
	
	
	
	
	
	private void parseHMCFiles(String baseDir) {
		if (managerType != M_HMC)
			return;
		
		BufferedReader br;
		String values[]=null;
		String line;
		int i;
		String s;
		Vector<String> fixes = new Vector<String>();
		
		
		hmc = new GenericData();
		
		// Identify BIOS value
		try {			
			br = new BufferedReader(new FileReader(baseDir + lshmcb),1024*1024);
			// only first line is meaningful
			if ( (line=br.readLine())!= null ) {
				i = line.indexOf('=');
				if (i>=0) {
					s = line.substring(i+1);
					values = new String[1];
					values[0]=s;
					hmc.add("bios", values);
				}			
			}			
		} 
		catch (IOException ioe) {}
		
		// Identify HMC server model and serial value
		try {			
			br = new BufferedReader(new FileReader(baseDir + lshmcv),1024*1024);		
			while ( (line=br.readLine())!= null ) {		
				// Only *TM and *SE are meaningful
				if (line.startsWith("*TM")) {
					s = line.substring(4);
					values = new String[1];
					values[0]=s;
					hmc.add("model", values);
				}
				if (line.startsWith("*SE")) {
					s = line.substring(4);
					values = new String[1];
					values[0]=s;
					hmc.add("serial", values);
				}	
			}					
		} 
		catch (IOException ioe) {}
		
		// Identify HMC software setup
		try {			
			br = new BufferedReader(new FileReader(baseDir + lshmcV),1024*1024);		
			while ( (line=br.readLine())!= null ) {	
				if ( (i=line.indexOf("Version:"))>=0 ) {
					s = line.substring(i+9);
					values = new String[1];
					values[0]=s;
					hmc.add("version", values);
				}
				if ( (i=line.indexOf("Release:"))>=0 ) {
					s = line.substring(i+9);
					values = new String[1];
					values[0]=s;
					hmc.add("release", values);
				}
				if ( (i=line.indexOf("Service Pack:"))>=0 ) {
					s = line.substring(i+14);
					values = new String[1];
					values[0]=s;
					hmc.add("sp", values);
				}
				if ( (i=line.indexOf("HMC Build level"))>=0 ) {
					s = line.substring(i+16);
					values = new String[1];
					values[0]=s;
					hmc.add("build_level", values);
				}
				if ( (i=line.indexOf("base_version"))>=0 ) {
					s = line.substring(i+13);
					values = new String[1];
					values[0]=s;
					hmc.add("base_version", values);
				}
				
				if (line.startsWith("MH")) {
					fixes.add(line);
				}
			}	
			
			if (fixes.size()>0) {
				values = new String[fixes.size()];
				for (i=0; i<fixes.size(); i++)
					values[i] = fixes.elementAt(i);
				hmc.add("fixes",values);
			}
		} 
		catch (IOException ioe) {}
		
		
		try {			
			br = new BufferedReader(new FileReader(baseDir + lshmcn),1024*1024);
			
			DataParser dp;
			String names[];
			
			while ( (line=br.readLine())!= null ) {
				
				dp = new DataParser(line);
				names = dp.getNames();
				
				// Skip line if no variable have been identified
				if (names==null || names.length==0 || names[0]==null)
					continue;
				
				for (i=0; i<names.length; i++) {
					hmc.add(names[i], dp.getStringValue(names[i]));					
				}				
			}			
		} 
		catch (IOException ioe) {}

		
	}
	
	
	
	private void parseDLPAR(String baseDir) {
		
		BufferedReader br;
		String line;
		String lparId;
		String serial;
		String ip;
		String active;
		String os;
		int pos1, pos2;
		GenericData gd;
		String values[];

	
		try {			
			br = new BufferedReader(new FileReader(baseDir + lspartition),1024*1024);
			
			while ( (line=br.readLine())!= null ) {
				pos1 = line.indexOf('<', 1)+1;
				pos2 = line.indexOf('*', pos1);
				lparId = line.substring(pos1, pos2);
				
				pos1 = line.indexOf('*', pos2+1)+1;
				pos2 = line.indexOf(',', pos1);
				serial = line.substring(pos1, pos2);
				
				pos1 = line.indexOf(',', pos2+1)+1;
				pos2 = line.indexOf('>', pos1);
				ip = line.substring(pos1, pos2).trim();
				
				line=br.readLine();
				pos1 = line.indexOf('<')+1;
				pos2 = line.indexOf('>', pos1);
				active = line.substring(pos1, pos2);
				
				pos1 = line.indexOf('<', pos2)+1;
				pos2 = line.indexOf(',', pos1);
				os = line.substring(pos1, pos2);
				pos1 = pos2+1;
				pos2 = line.indexOf(',', pos1);
				os = os + " " + line.substring(pos1, pos2).trim();
				pos1 = pos2+1;
				pos2 = line.indexOf('>', pos1);
				os = os + " " + line.substring(pos1, pos2).trim();
				
				
				gd = getLpar(serial, lparId);
				if (gd!=null) {
					values = new String[1];
					values[0]=ip;
					gd.add("rmc_ipaddr", values);
					values = new String[1];
					if (active.equals("1"))
						values[0]="active";
					else
						values[0]="off";
					gd.add("rmc_state", values);	
					
					// Update OS if you have found a value
					if (os.length()>2) {
						values = gd.getVarValues("os_version");
						// If it starts with VIOS it has already been asked to VIOS!!!!
						if (values==null || !values[0].startsWith("VIOS")) {
							values = new String[1];
							values[0]=os;
							gd.add("os_version", values);
						}
					}
						
				}
			}
		} 
		catch (IOException ioe) {}	
	}
	
	
	
	private GenericData getLpar(String serial, String lparID) {
		int i,j;
		GenericData gd[];
		
		for (i=0; i<managedSystem.length; i++) {
			if (managedSystem[i].getVarValues("serial_num")[0].equals(serial)) {
				gd = managedSystem[i].getObjects(CONFIG_LPAR);
				for (j=0; j<gd.length; j++) {
					if (gd[j].getVarValues("lpar_id")[0].equals(lparID))
						return gd[j];
				}
			}
		}
		
		return null;
	}
	
	
	
	private void parseData(String hmc, String baseDir) {
		int i,j;
		
		// Load scanner parameters
		loadScannerParams(baseDir);
		
		// Load system data into parent data structure (already done if HMC was called)
		if (onlyReadFile) {
			loadSysConfigData(hmc, baseDir);
			if (managerType == M_HMC)
				parseEntPool(baseDir);
		}
		
		if (managedSystem == null) {
			System.out.println(" no managed systems detected!");
			return;
		}
	
		for (i=0; i<managedSystem.length; i++) {
			load_data(managedSystem[i], baseDir, procSysData,	 PROC);
			load_data(managedSystem[i], baseDir, memSysData,	 MEM);
			load_data(managedSystem[i], baseDir, slotSysData, 	 SLOT);
			load_data(managedSystem[i], baseDir, procLparData, 	 PROC_LPAR);
			load_data(managedSystem[i], baseDir, procPoolData, 	 PROC_POOL);
			load_data(managedSystem[i], baseDir, memLparData, 	 MEM_LPAR);
			load_data(managedSystem[i], baseDir, memPoolData, 	 MEM_POOL);
			load_data(managedSystem[i], baseDir, lparConfigData, CONFIG_LPAR);
			load_data(managedSystem[i], baseDir, lparProfiles, 	 PROFILES);
			load_data(managedSystem[i], baseDir, vswitchData, 	 VSWITCH);
			load_data(managedSystem[i], baseDir, vethData, 		 VETH);
			load_data(managedSystem[i], baseDir, vscsiData, 	 VSCSI);
			load_data(managedSystem[i], baseDir, vfcData, 		 VFC);
			load_data(managedSystem[i], baseDir, lslicSyspower,  SYSPOWERLIC);
			load_data(managedSystem[i], baseDir, slotchildren,   IOSLOTCHILDREN);
			
			load_data(managedSystem[i], baseDir, lscod_bill_proc,   	LSCOD_BILL_PROC);
			load_data(managedSystem[i], baseDir, lscod_bill_mem,   		LSCOD_BILL_MEM);
			load_data(managedSystem[i], baseDir, lscod_cap_proc_onoff,  LSCOD_CAP_PROC_ONOFF);
			load_data(managedSystem[i], baseDir, lscod_cap_mem_onoff,  	LSCOD_CAP_MEM_ONOFF);
			load_data(managedSystem[i], baseDir, lscod_hist,  			LSCOD_HIST);
			
			load_npiv_data(managedSystem[i], baseDir);
			
			System.out.print(".");
		}
		
		if (managerType == M_HMC)
			parseEntPoolSystem(baseDir);
		
		if (managerType == M_HMC)
			parseHMCFiles(baseDir);
		
		//parseDLPAR(baseDir);
		
		if (produceStatistics && managerType==M_IVM) {
			//produceStatistics=false;
		}
		
		if (produceStatistics) {
			
			
		
			// Load perf data
			//load_saved_stats(baseDir);
			
			int num=0;
			GenericData gd[];
			for (i=0; i<managedSystem.length; i++) {
				gd = managedSystem[i].getObjects(CONFIG_LPAR);
				if (gd!=null)
					num += gd.length;
			}
			lparNames = new String[num];
			lparID = new String[num];
			msCoreConfig = new DataManager[managedSystem.length];
			msCoreAvail = new DataManager[managedSystem.length];
			msCoreUsed = new DataManager[managedSystem.length];
			msMemConfig = new DataManager[managedSystem.length];
			msMemAvail = new DataManager[managedSystem.length];
			lparEnt = new DataManager[lparNames.length];
			lparVP = new DataManager[lparNames.length];
			lparPC = new DataManager[lparNames.length];
			lparStatus = new NewLparStatus[lparNames.length];
			
			num=0;
			for (i=0; i<managedSystem.length; i++) {
				gd = managedSystem[i].getObjects(CONFIG_LPAR);
				if (gd==null)
					continue;
				for (j=0; j<gd.length; j++) {
					lparNames[num] = gd[j].getVarValues("name")[0];
					lparID[num] = gd[j].getVarValues("lpar_id")[0];
					num++;
				}
			}
			
			for (i=0; i<managedSystem.length; i++) {
				msCoreConfig[i] = new DataManager(scannerDate);
				msCoreAvail[i] = new DataManager(scannerDate);
				msCoreUsed[i] = new DataManager(scannerDate);
				msMemConfig[i] = new DataManager(scannerDate);
				msMemAvail[i] = new DataManager(scannerDate);				
			}
			for (i=0; i<lparNames.length; i++) {
				lparEnt[i] = new DataManager(scannerDate);
				lparVP[i] = new DataManager(scannerDate);
				lparPC[i] = new DataManager(scannerDate);	
				lparStatus[i] = new NewLparStatus(scannerDate);
			}
			
			
			load_all_stats(baseDir);
			
			
			// DELETE
			
			//for (i=0; i<managedSystem.length; i++)
			//	load_stats(i, baseDir);
			//save_stats(baseDir);
			
			// Sort names
			sortManagedData();
			sortLparData();
			
			// Create summary
			//compute_summary();
		}
		
		System.out.println(" DONE");
	}
	
	
	
	private float sysdata_average(int system, int item, int from, int to) {
		int i; 
		double sum, num;
		
		sum=num=0;
		for (i=from; i<=to; i++)
			if (goodSystemData[system][i]) {
				sum += managedSystemData[system][i][item];
				num++;
			}
		if (num==0)
			return -1;
		return (float)(sum/num);		
	}
	
	
	private float lpardata_average(int lpar, int item, int from, int to) {
		int i; 
		double sum, num;
		
		sum=num=0;
		for (i=from; i<=to; i++)
			if (goodLparData[lpar][i]) {
				sum += lparData[lpar][i][item];
				num++;
			}
		if (num==0)
			return -1;
		return (float)(sum/num);		
	}
	
	
	
	private void compute_summary() {
		int i,j;
		
		managedSystemDataSummary = new float[managedSystem.length][NUM_SUMMARY][NUM_MS];
		
		for (i=0; i<managedSystem.length; i++) {
			for (j=0; j<NUM_MS; j++) {			
				managedSystemDataSummary[i][WEEK1][j] = sysdata_average(i, j, 0, 6);
				managedSystemDataSummary[i][WEEK2][j] = sysdata_average(i, j, 7, 13);
				managedSystemDataSummary[i][WEEK3][j] = sysdata_average(i, j, 14, 20);
				managedSystemDataSummary[i][WEEK4][j] = sysdata_average(i, j, 21, 27);
				managedSystemDataSummary[i][MONTH1][j] = sysdata_average(i, j, 0, 29);
				managedSystemDataSummary[i][MONTH2][j] = sysdata_average(i, j, 30, 59);
				managedSystemDataSummary[i][MONTH3][j] = sysdata_average(i, j, 60, 89);
				managedSystemDataSummary[i][MONTH4][j] = sysdata_average(i, j, 90, 119);
				managedSystemDataSummary[i][MONTH5][j] = sysdata_average(i, j, 120, 149);
				managedSystemDataSummary[i][MONTH6][j] = sysdata_average(i, j, 150, 179);
				managedSystemDataSummary[i][MONTH7][j] = sysdata_average(i, j, 180, 209);
				managedSystemDataSummary[i][MONTH8][j] = sysdata_average(i, j, 210, 239);
				managedSystemDataSummary[i][MONTH9][j] = sysdata_average(i, j, 240, 269);
				managedSystemDataSummary[i][MONTH10][j] = sysdata_average(i, j, 270, 299);
				managedSystemDataSummary[i][MONTH11][j] = sysdata_average(i, j, 300, 329);
				managedSystemDataSummary[i][MONTH12][j] = sysdata_average(i, j, 330, 359);
			}
		}	
		
		
		lparDataSummary = new float[lparData.length][NUM_SUMMARY][NUM_LPAR];
		
		for (i=0; i<lparData.length; i++) {
			for (j=0; j<NUM_LPAR; j++) {			
				lparDataSummary[i][WEEK1][j] = lpardata_average(i, j, 0, 6);
				lparDataSummary[i][WEEK2][j] = lpardata_average(i, j, 7, 13);
				lparDataSummary[i][WEEK3][j] = lpardata_average(i, j, 14, 20);
				lparDataSummary[i][WEEK4][j] = lpardata_average(i, j, 21, 27);
				lparDataSummary[i][MONTH1][j] = lpardata_average(i, j, 0, 29);
				lparDataSummary[i][MONTH2][j] = lpardata_average(i, j, 30, 59);
				lparDataSummary[i][MONTH3][j] = lpardata_average(i, j, 60, 89);
				lparDataSummary[i][MONTH4][j] = lpardata_average(i, j, 90, 119);
				lparDataSummary[i][MONTH5][j] = lpardata_average(i, j, 120, 149);
				lparDataSummary[i][MONTH6][j] = lpardata_average(i, j, 150, 179);
				lparDataSummary[i][MONTH7][j] = lpardata_average(i, j, 180, 209);
				lparDataSummary[i][MONTH8][j] = lpardata_average(i, j, 210, 239);
				lparDataSummary[i][MONTH9][j] = lpardata_average(i, j, 240, 269);
				lparDataSummary[i][MONTH10][j] = lpardata_average(i, j, 270, 299);
				lparDataSummary[i][MONTH11][j] = lpardata_average(i, j, 300, 329);
				lparDataSummary[i][MONTH12][j] = lpardata_average(i, j, 330, 359);
			}
		}
	}
	
	
	
	private int getLparIdByName(String name) {
		int i=0;
		
		while (i<lparNames.length && !lparNames[i].equals(name))
			i++;
		
		if (i<lparNames.length)
			return i;
		
		return -1;
	}
	
	
	private int getLparIdById(String name) {
		int i=0;
		
		while (i<lparID.length && !lparID[i].equals(name))
			i++;
		
		if (i<lparID.length)
			return i;
		
		return -1;
	}
	
	
	
	
	/*
	 * Prepare data structure and load saved data if available
	 */
	private void load_saved_stats(String baseDir) {
		BufferedReader br;
		String names[]=null;
		String line;
		DataParser dp;
		int i,j,k;
		//GregorianCalendar today = new GregorianCalendar();
		GregorianCalendar today = scannerDate;
		GregorianCalendar sample;
		int day;
		String s;
		long diffMillis;
		int num;
		GenericData gd[];
		

		
		managedSystemData = new float[managedSystem.length][TIME_PERIOD][NUM_MS];
		for (i=0; i<managedSystem.length; i++)
			for (j=0; j<TIME_PERIOD; j++)
				for (k=0; k<NUM_MS; k++)
					managedSystemData[i][j][k]=0;
			
		goodSystemData = new boolean[managedSystem.length][TIME_PERIOD];
		for (i=0; i<managedSystem.length; i++)
			for (j=0; j<TIME_PERIOD; j++)
				goodSystemData[i][j]=false;
		
		

		num=0;
		for (i=0; i<managedSystem.length; i++) {
			gd = managedSystem[i].getObjects(CONFIG_LPAR);
			if (gd!=null)
				num += gd.length;
		}
		lparNames = new String[num];
		
		num=0;
		for (i=0; i<managedSystem.length; i++) {
			gd = managedSystem[i].getObjects(CONFIG_LPAR);
			if (gd==null)
				continue;
			for (j=0; j<gd.length; j++)
				lparNames[num++] = gd[j].getVarValues("name")[0];
		}
		
		lparData = new float[lparNames.length][TIME_PERIOD][NUM_LPAR];
		for (i=0; i<lparData.length; i++)
			for (j=0; j<TIME_PERIOD; j++)
				for (k=0; k<NUM_LPAR; k++)
					lparData[i][j][k]=0;
			
		goodLparData = new boolean[lparData.length][TIME_PERIOD];
		for (i=0; i<goodLparData.length; i++)
			for (j=0; j<TIME_PERIOD; j++)
				goodLparData[i][j]=false;
		
		lpar_cycles = new BigDecimal[lparNames.length][TIME_PERIOD][NUM_LPAR_CYCLES];
		for (i=0; i<lparNames.length; i++)
			for (j=0; j<TIME_PERIOD; j++)
				for (k=0; k<NUM_LPAR_CYCLES; k++)
					lpar_cycles[i][j][k] = new BigDecimal(0);

		
		/*
		// Load previously saved data
		try {			

			br = new BufferedReader( 
					new InputStreamReader(
							new GZIPInputStream(
									new FileInputStream(baseDir + systemPerf)
							)
					) ,
					1024*1024);
			
			while ( (line=br.readLine())!= null ) {
				
				dp = new DataParser(line);
				names = dp.getNames();
				
				// Skip line if no variable have been identified
				if (names==null || names.length==0 || names[0]==null)
					continue;
				
				s = dp.getStringValue("time")[0];
				
				// 0123456789
				// 2012/12/31
				sample = new GregorianCalendar(		Integer.parseInt(s.substring(0,4)),
													Integer.parseInt(s.substring(5,7))-1,
													Integer.parseInt(s.substring(8)));
				diffMillis = today.getTimeInMillis()-sample.getTimeInMillis();
				day = (int)(diffMillis/(24*60*60*1000));
				
				if (day>=TIME_PERIOD)
					continue;
				
				s = dp.getStringValue("ms")[0];
				for (i=0; i<managedSystem.length; i++)
					if (managedSystem[i].getVarValues("name")[0].equals(s))
						break;
				if (i>=managedSystem.length)
					return;

				managedSystemData[i][day][CONFIGURABLE_POOL_PROC_UNITS] = Float.parseFloat(dp.getStringValue("configurable_pool_proc_units")[0]);
				managedSystemData[i][day][CURR_AVAIL_POOL_PROC_UNITS] = Float.parseFloat(dp.getStringValue("curr_avail_pool_proc_units")[0]);
				managedSystemData[i][day][BORROWED_POOL_PROC_UNITS] = Float.parseFloat(dp.getStringValue("borrowed_pool_proc_units")[0]);
				managedSystemData[i][day][CONFIGURABLE_SYS_MEM] = Float.parseFloat(dp.getStringValue("configurable_sys_mem")[0]);
				managedSystemData[i][day][USED_POOL] = Float.parseFloat(dp.getStringValue("used_pool")[0]);
				managedSystemData[i][day][AVAIL_SYS_MEM] = Float.parseFloat(dp.getStringValue("curr_avail_sys_mem")[0]);
				goodSystemData[i][day]=true;	
			}			
		} 
		catch (IOException ioe) {	

		}
		*/
		
	}
	
	
	
	/*
	 * Save loaded data for future use
	 */
	private void save_stats(String baseDir) {
		int i,j;
		//GregorianCalendar today = new GregorianCalendar();
		GregorianCalendar today = scannerDate;
		GregorianCalendar sample, last;
		PrintWriter 			writer = null;
		long 					millis	= today.getTime().getTime();
		long 					delta;
		BigDecimal				bi1, bi2, bi3;
		
		
		// Create LPAR data from collected cycles
		for (i=0; i< lparData.length; i++)
			for (j=0; j<TIME_PERIOD; j++) {
				bi1 = lpar_cycles[i][j][LPAR_CYCLES];
				bi2 = lpar_cycles[i][j][LPAR_TIME_CYCLES];
				if (bi2.compareTo(zero)!=0) {
					bi3=bi1.divide(bi2,new MathContext(5));
					lparData[i][j][USED_CPU] = bi3.floatValue();
					goodLparData[i][j]=true;
				} else
					lparData[i][j][USED_CPU] = -1;
			}	
		
		/*
		sample = new GregorianCalendar();
		last = new GregorianCalendar(2000,0,1);  
		
		try {
			writer = new PrintWriter(new GZIPOutputStream(new FileOutputStream(baseDir + systemPerf)));
			
			for (i=0; i<managedSystem.length; i++)
				for (j=0; j<TIME_PERIOD; j++) {
					if (!goodSystemData[i][j])
						continue;
					delta = 1000l*60*60*24*j;
					sample.setTime(new Date(millis-delta));
					//sample.setTimeInMillis(millis - j*1000*60*60*24);
					if (sample.after(last))
						last = (GregorianCalendar)sample.clone();
					
					writer.print("time=" +	sample.get(Calendar.YEAR) + "/");
					if ( (sample.get(Calendar.MONTH)+1) < 10)
						writer.print("0");
					writer.print((sample.get(Calendar.MONTH)+1) + "/");
					if ( sample.get(Calendar.DAY_OF_MONTH) < 10)
						writer.print("0");
					writer.print(sample.get(Calendar.DAY_OF_MONTH) + ",");
					
					writer.println(	"ms=" +
									managedSystem[i].getVarValues("name")[0] + "," +
									"configurable_pool_proc_units=" + 
									managedSystemData[i][j][CONFIGURABLE_POOL_PROC_UNITS] + "," +
									"curr_avail_pool_proc_units=" + 
									managedSystemData[i][j][CURR_AVAIL_POOL_PROC_UNITS] + "," +
									"borrowed_pool_proc_units=" + 
									managedSystemData[i][j][BORROWED_POOL_PROC_UNITS] + "," +
									"configurable_sys_mem=" + 
									managedSystemData[i][j][CONFIGURABLE_SYS_MEM] + "," + 
									"used_pool=" +
									managedSystemData[i][j][USED_POOL] + "," + 
									"curr_avail_sys_mem=" +
									managedSystemData[i][j][AVAIL_SYS_MEM]);			
				}			
			writer.close();			
			
		} catch (Exception e) {
			if (writer!=null)
				writer.close();
		}	
		
		try {
			writer = new PrintWriter(new BufferedWriter(new FileWriter(baseDir + lastPerfDate)));
			writer.print(last.get(Calendar.YEAR) + "/");
			if (last.get(Calendar.MONTH)+1<10)
				writer.print("0");
			writer.print(		(last.get(Calendar.MONTH)+1) + "/");
			if (last.get(Calendar.DAY_OF_MONTH)<10)
				writer.print("0");
			writer.println(		last.get(Calendar.DAY_OF_MONTH));
			writer.close();
		} catch (Exception e) {
			if (writer!=null)
				writer.close();
		}
		*/
	}
	
	
	
	private void load_stats(int num, String baseDir) {
		String sysName;
		BufferedReader br;
		String names[]=null;
		String line;
		DataParser dp;
		String resource_type;
		//GregorianCalendar today = new GregorianCalendar();
		GregorianCalendar today = scannerDate;
		GregorianCalendar sample;
		GregorianCalendar prev_sample = null;
		int day = -1;
		String s;
		long diffMillis;
		
		int prevDay = -1;
		BigDecimal	currPoolData[] = new BigDecimal[NUM_POOLDATA];
		BigDecimal	prevPoolData[] = new BigDecimal[NUM_POOLDATA];
		BigDecimal  currLparData[][] = new BigDecimal[lparNames.length][NUM_LPAR_CYCLES];
		BigDecimal  prevLparData[][] = new BigDecimal[lparNames.length][NUM_LPAR_CYCLES];
		int i,j;
		BigDecimal bi1, bi2, bi3;
		
		
		sysName = managedSystem[num].getVarValues("name")[0];
			
		// Load new data from new file
		try {			

			br = new BufferedReader( 
					new InputStreamReader(
							new GZIPInputStream(
									new FileInputStream(baseDir + sysName + "_" + lslparutilStats)
							)
					) ,
					1024*1024);
			
			while ( (line=br.readLine())!= null ) {
				
				// Skip line if no data is returned
				if (line.startsWith("HSC") || line.startsWith("The managed system") || 
						line.startsWith("No results were found") ||
								line.contains("state=Standby") )
					break;
				
				dp = new DataParser(line);
				names = dp.getNames();
				
				// Skip line if no variable have been identified
				if (names==null || names.length==0 || names[0]==null)
					continue;
				
				resource_type = dp.getStringValue("resource_type")[0];
				s = dp.getStringValue("time")[0];
				sample = new GregorianCalendar(		Integer.parseInt(s.substring(6,10)),
													Integer.parseInt(s.substring(0,2))-1,
													Integer.parseInt(s.substring(3,5)));
				if (prev_sample == null)
					prev_sample = sample;
				
				if (!sample.equals(prev_sample)) {
					prevDay=day;
					prev_sample = sample;
					for (i=0; i<currLparData.length; i++)
						for (j=0; j<NUM_LPAR_CYCLES; j++) {
							prevLparData[i][j]=currLparData[i][j];
							currLparData[i][j]=null;
						}
				}	
								
				diffMillis = today.getTimeInMillis()-sample.getTimeInMillis();
				day = (int)(diffMillis/(24*60*60*1000));
				
				
				
				if (day>=TIME_PERIOD)
					continue;
				
				
				
				if (resource_type.equals("sys")) {
					// If not operating the data is void
					if (!dp.getStringValue("state")[0].equals("Operating"))
						continue;
					managedSystemData[num][day][CONFIGURABLE_SYS_MEM] = Float.parseFloat(dp.getStringValue("configurable_sys_mem")[0]);
					managedSystemData[num][day][AVAIL_SYS_MEM] = Float.parseFloat(dp.getStringValue("curr_avail_sys_mem")[0]);
					goodSystemData[num][day]=true;
					continue;
				}
				
				if (resource_type.equals("pool")) {
					managedSystemData[num][day][CONFIGURABLE_POOL_PROC_UNITS] = Float.parseFloat(dp.getStringValue("configurable_pool_proc_units")[0]);
					managedSystemData[num][day][CURR_AVAIL_POOL_PROC_UNITS] = Float.parseFloat(dp.getStringValue("curr_avail_pool_proc_units")[0]);
					managedSystemData[num][day][BORROWED_POOL_PROC_UNITS] = Float.parseFloat(dp.getStringValue("borrowed_pool_proc_units")[0]);
					
					currPoolData[TOTAL_POOL_CYCLES] = new BigDecimal(dp.getStringValue("total_pool_cycles")[0]);
					currPoolData[UTILIZED_POOL_CYCLES] = new BigDecimal(dp.getStringValue("utilized_pool_cycles")[0]);
					currPoolData[POOL_TIME_CYCLES] = new BigDecimal(dp.getStringValue("time_cycles")[0]);
					
					if (prevDay>=0) {
						// Not the very first sample		
						bi1 = prevPoolData[UTILIZED_POOL_CYCLES].subtract(currPoolData[UTILIZED_POOL_CYCLES]);
						bi2 = prevPoolData[POOL_TIME_CYCLES].subtract(currPoolData[POOL_TIME_CYCLES]);
						bi3 = bi1.divide(bi2, new MathContext(5));
						managedSystemData[num][prevDay][USED_POOL] = bi3.floatValue();
					}
					
					for (i=0; i<NUM_POOLDATA; i++) 
						prevPoolData[i]=currPoolData[i];
					//prevDay = day;
					
					continue;
				}
				
				if (resource_type.equals("lpar")) {
					
					// Skip LPARs that are not shared
					String mode = dp.getStringValue("curr_proc_mode")[0];
					if (!mode.equals("shared"))
						continue;
					
					int id = getLparIdByName(dp.getStringValue("lpar_name")[0]);
					if (id < 0) {
						// LPAR has been deleted
						continue;
					}
					
					if (!dp.getStringValue("curr_proc_mode")[0].equals("shared"))
						continue;
					
					lparData[id][day][CURR_PROC_UNITS] = Float.parseFloat(dp.getStringValue("curr_proc_units")[0]);
					lparData[id][day][CURR_PROCS] = Float.parseFloat(dp.getStringValue("curr_procs")[0]);
					
					bi1 = new BigDecimal(dp.getStringValue("capped_cycles")[0]);
					bi2 = new BigDecimal(dp.getStringValue("uncapped_cycles")[0]);
					currLparData[id][LPAR_CYCLES] = bi1.add(bi2);
					bi1 = new BigDecimal(dp.getStringValue("time_cycles")[0]);
					currLparData[id][LPAR_TIME_CYCLES] = bi1;
					
					if (prevDay>=0 && prevLparData[id][LPAR_CYCLES]!=null) {
						// Not the very first sample		
						bi1 = prevLparData[id][LPAR_CYCLES].subtract(currLparData[id][LPAR_CYCLES]);
						bi2 = prevLparData[id][LPAR_TIME_CYCLES].subtract(currLparData[id][LPAR_TIME_CYCLES]);
						lpar_cycles[id][prevDay][LPAR_CYCLES]=lpar_cycles[id][prevDay][LPAR_CYCLES].add(bi1);
						lpar_cycles[id][prevDay][LPAR_TIME_CYCLES]=lpar_cycles[id][prevDay][LPAR_TIME_CYCLES].add(bi2);
					}
										
					continue;
				}
				 	
			}			
		} 
		catch (IOException ioe) {	
			System.out.println("Loader.load_stats: IOException");
			System.out.println(ioe);
		}
				
	}
	
	
	/*
	 * Load hourly and daily stats
	 */
	private void load_all_stats(String baseDir) {
		
		int i;
		String sysName;
		
		for (i=0; i<managedSystem.length; i++) {
			sysName = managedSystem[i].getVarValues("name")[0];
			load_single_stats(i, baseDir + sysName + "_" + lslparutilHourlyStats, HOURLY);	
		}

		
		for (i=0; i<managedSystem.length; i++) {
			sysName = managedSystem[i].getVarValues("name")[0];
			load_single_stats(i, baseDir + sysName + "_" + lslparutilDailyStats, DAILY);
		}		
	}
	
	
	
	
	/*
	 * Identify proc pool by managed system name and pool name.
	 * If it does not exists, create the DataManaged object
	 */
	private int getProcPoolId(String ms, String poolName) {
		String s = poolName + "\n" + ms;
		
		if (procPoolName == null) {
			// Very first subpool
			procPoolName = new String[1];
			procPoolName[0]=s;
			procPoolConfig = new DataManager[1];
			procPoolConfig[0] = new DataManager(scannerDate);
			procPoolUsed = new DataManager[1];
			procPoolUsed[0] = new DataManager(scannerDate);
			return 0;
		}
		
		int n = 0;
		while (n<procPoolName.length && !s.equals(procPoolName[n]))
			n++;
		
		if (n==procPoolName.length) {
			// New Proc Pool
			int i;
			String oldNames[] = procPoolName;
			DataManager oldPoolConfig[] = procPoolConfig;
			DataManager oldPoolUsed[] = procPoolUsed;
			
			procPoolName = new String[procPoolName.length+1];
			procPoolConfig = new DataManager[procPoolConfig.length+1];
			procPoolUsed = new DataManager[procPoolUsed.length+1];
			
			for (i=0; i<oldNames.length; i++) {
				procPoolName[i] = oldNames[i];
				procPoolConfig[i] = oldPoolConfig[i];
				procPoolUsed[i] = oldPoolUsed[i];
			}
			
			procPoolName[n] = s;
			procPoolConfig[n] = new DataManager(scannerDate);
			procPoolUsed[n] = new DataManager(scannerDate);
			
			return n;		
		}
		
		// existing proc Pool
		return n;
	}
	
	
	
	/*
	 * Load single file data
	 */
	private void load_single_stats(int num, String fileName, byte type) {
		
		BufferedReader br;
		String names[]=null;
		String line;
		DataParser dp;
		String resource_type;
		String event_type;
		GregorianCalendar sample;
		String s;
		BigDecimal	currPoolData[] = new BigDecimal[NUM_POOLDATA];
		BigDecimal	prevPoolData[] = new BigDecimal[NUM_POOLDATA];
		GregorianCalendar prevPoolSample = null;
		BigDecimal  currLparData[][] = new BigDecimal[lparNames.length][NUM_LPAR_CYCLES];
		BigDecimal  prevLparData[][] = new BigDecimal[lparNames.length][NUM_LPAR_CYCLES];
		GregorianCalendar prevLparSample[] = new GregorianCalendar[lparNames.length];
		
		final int MAX_POOLS = 256;
		BigDecimal	currProcPoolData[][] = new BigDecimal[MAX_POOLS][NUM_POOLDATA];
		BigDecimal	prevProcPoolData[][] = new BigDecimal[MAX_POOLS][NUM_POOLDATA];
		String		currProcPoolName[] = new String[MAX_POOLS];
		GregorianCalendar prevProcPoolSample[] = new GregorianCalendar[MAX_POOLS];
		
		String ss[];
		boolean cap;
		String poolName;
		
		
		
		int i,j;
		BigDecimal bi1, bi2, bi3, bi4;
		
			
		// Load new data from hourly file
		try {			

			br = new BufferedReader( 
					new InputStreamReader(
							new GZIPInputStream(
									new FileInputStream(fileName)
							)
					) ,
					1024*1024);
			
			while ( (line=br.readLine())!= null ) {
				
				// Skip line if no data is returned
				if (line.startsWith("HSC") || line.startsWith("The managed system") || 
						line.startsWith("No results were found")  )
					break;
				
				if (line.contains("state=Standby"))
					continue;
				
				dp = new DataParser(line);
				names = dp.getNames();
				
				// Skip line if no variable have been identified
				if (names==null || names.length==0 || names[0]==null)
					continue;
				
				event_type = dp.getStringValue("event_type")[0];
				if (event_type==null || !event_type.equals("sample"))
					continue;
				
				resource_type = dp.getStringValue("resource_type")[0];
				s = dp.getStringValue("time")[0];
				// 04/03/2013 23:00:02
				// 0123456789012345678901234567890
				// 0         1         2         3
				sample = new GregorianCalendar(		Integer.parseInt(s.substring(6,10)),
													Integer.parseInt(s.substring(0,2))-1,
													Integer.parseInt(s.substring(3,5)),
													Integer.parseInt(s.substring(11,13)),
													Integer.parseInt(s.substring(14, 16)));
				
							
				// Handle current record and store data in current data structures
				
				if ( resource_type.equals("sys") ) {
					
					// If not operating the data is void
					ss = dp.getStringValue("state");
					if (ss==null) {
						// New HMC has changed string!!!!
						ss = dp.getStringValue("primary_state");
					}
					if (ss==null) {
						System.out.println("load_single_stats: sys sample with missing state? Aborting.");
						System.exit(1);
					}
						
					// If not operating the data is void
					if ( !ss[0].equals("Operating") && !ss[0].equals("Started")) 
						continue;
					
					if (type==HOURLY) {
						msMemConfig[num].addHourData(sample, 
								Float.parseFloat(dp.getStringValue("configurable_sys_mem")[0])/1024);
						msMemAvail[num].addHourData(sample, 
								Float.parseFloat(dp.getStringValue("curr_avail_sys_mem")[0])/1024);
					} else if (type==DAILY) {
						msMemConfig[num].addDayData(sample, 
								Float.parseFloat(dp.getStringValue("configurable_sys_mem")[0])/1024);
						msMemAvail[num].addDayData(sample, 
								Float.parseFloat(dp.getStringValue("curr_avail_sys_mem")[0])/1024);
					} 
					
					continue;
					
				} 
				
				
				if (resource_type.equals("pool")) {
					
					// Handle current values
					
					if (type==HOURLY) {
						msCoreConfig[num].addHourData(sample,
								Float.parseFloat(dp.getStringValue("configurable_pool_proc_units")[0]) +
								Float.parseFloat(dp.getStringValue("borrowed_pool_proc_units")[0]));
						msCoreAvail[num].addHourData(sample, 
								Float.parseFloat(dp.getStringValue("curr_avail_pool_proc_units")[0]));
					} else if (type==DAILY) {
						msCoreConfig[num].addDayData(sample,
								Float.parseFloat(dp.getStringValue("configurable_pool_proc_units")[0]) +
								Float.parseFloat(dp.getStringValue("borrowed_pool_proc_units")[0]));
						msCoreAvail[num].addDayData(sample, 
								Float.parseFloat(dp.getStringValue("curr_avail_pool_proc_units")[0]));
					}
				
					currPoolData[TOTAL_POOL_CYCLES] = new BigDecimal(dp.getStringValue("total_pool_cycles")[0]);
					currPoolData[UTILIZED_POOL_CYCLES] = new BigDecimal(dp.getStringValue("utilized_pool_cycles")[0]);
					currPoolData[POOL_TIME_CYCLES] = new BigDecimal(dp.getStringValue("time_cycles")[0]);
					
					
					// Handle deltas. Skip if time cycles wrap
					if ( prevPoolSample != null &&
							prevPoolData[UTILIZED_POOL_CYCLES]!=null &&
							prevPoolData[POOL_TIME_CYCLES]!=null &&
							currPoolData[UTILIZED_POOL_CYCLES]!=null &&
							currPoolData[POOL_TIME_CYCLES]!=null &&
							prevPoolData[POOL_TIME_CYCLES].compareTo(currPoolData[POOL_TIME_CYCLES])>0 /* &&
							prevPoolData[UTILIZED_POOL_CYCLES].compareTo(currPoolData[UTILIZED_POOL_CYCLES])>0 */) {
						
						bi1 = prevPoolData[UTILIZED_POOL_CYCLES].subtract(currPoolData[UTILIZED_POOL_CYCLES]);
						bi2 = prevPoolData[POOL_TIME_CYCLES].subtract(currPoolData[POOL_TIME_CYCLES]);					
						bi3 = bi1.divide(bi2, new MathContext(5));
						//System.out.println(bi3.floatValue());
						
						if (type==HOURLY)
							msCoreUsed[num].addHourData(prevPoolSample, bi3.floatValue());
						else if (type==DAILY)
							msCoreUsed[num].addDayData(prevPoolSample, bi3.floatValue());
					}
					
					// Store data in previous record
					prevPoolSample = sample;
					prevPoolData[UTILIZED_POOL_CYCLES] = currPoolData[UTILIZED_POOL_CYCLES];
					prevPoolData[POOL_TIME_CYCLES] = currPoolData[POOL_TIME_CYCLES];
					
					continue;
					
					
					//System.out.println(currPoolData[UTILIZED_POOL_CYCLES].floatValue());
					
				} 
				
				if (resource_type.equals("procpool")) {
					
					int poolID = Integer.parseInt(dp.getStringValue("shared_proc_pool_id")[0]);
					
					// Skip pool 0
					if (poolID==0)
						continue;
					
					// Only keep latest pool name!
					if (currProcPoolName[poolID] == null)
						currProcPoolName[poolID] = dp.getStringValue("shared_proc_pool_name")[0];
					
					currProcPoolData[poolID][TOTAL_POOL_CYCLES] = new BigDecimal(dp.getStringValue("total_pool_cycles")[0]);
					currProcPoolData[poolID][UTILIZED_POOL_CYCLES] = new BigDecimal(dp.getStringValue("utilized_pool_cycles")[0]);
					currProcPoolData[poolID][POOL_TIME_CYCLES] = new BigDecimal(dp.getStringValue("time_cycles")[0]);
									
					
					// Handle deltas. Skip if time cycles wrap
					if (	prevProcPoolData[poolID][UTILIZED_POOL_CYCLES]!=null &&
							prevProcPoolData[poolID][POOL_TIME_CYCLES]!=null &&
							currProcPoolData[poolID][UTILIZED_POOL_CYCLES]!=null &&
							currProcPoolData[poolID][POOL_TIME_CYCLES]!=null &&
							prevProcPoolData[poolID][POOL_TIME_CYCLES].compareTo(currProcPoolData[poolID][POOL_TIME_CYCLES])>0 /* &&
							prevProcPoolData[poolID][UTILIZED_POOL_CYCLES].compareTo(currProcPoolData[poolID][UTILIZED_POOL_CYCLES])>0 */) {
						
						int id = getProcPoolId(managedSystem[num].getVarValues("name")[0], currProcPoolName[poolID]);
						
						bi1 = prevProcPoolData[poolID][TOTAL_POOL_CYCLES].subtract(currProcPoolData[poolID][TOTAL_POOL_CYCLES]);
						bi4 = prevProcPoolData[poolID][POOL_TIME_CYCLES].subtract(currProcPoolData[poolID][POOL_TIME_CYCLES]);
						bi2 = bi1.divide(bi4, new MathContext(5));	// POOL SIZE
						
						if (type==HOURLY)
							procPoolConfig[id].addHourData(prevProcPoolSample[poolID], bi2.floatValue()); // Add pool Size
						else if (type==DAILY)
							procPoolConfig[id].addDayData(prevProcPoolSample[poolID], bi2.floatValue()); // Add pool Size
						

						bi3 = prevProcPoolData[poolID][UTILIZED_POOL_CYCLES].subtract(currProcPoolData[poolID][UTILIZED_POOL_CYCLES]);
						
						/*
						// If no pool cycles are used, do not compare it with pool size (it may be zero!!)
						if (bi3.compareTo(new BigDecimal(0))==0)
							bi4 = bi3;
						else
							bi4 = bi3.divide(bi1, new MathContext(5)).multiply(bi2, new MathContext(5));
						procPoolUsed[id].add(prevProcPoolSample[poolID], bi4.floatValue());// Add pool used
						*/
						
						bi4 = bi3.divide(bi4, new MathContext(5));  // POOL USED
						
						if (type==HOURLY)
							procPoolUsed[id].addHourData(prevProcPoolSample[poolID], bi4.floatValue());// Add pool used
						else if (type==DAILY)
							procPoolUsed[id].addDayData(prevProcPoolSample[poolID], bi4.floatValue());// Add pool used
					}
					
					// Store data in previous record
					prevProcPoolSample[poolID] = sample;
					prevProcPoolData[poolID][UTILIZED_POOL_CYCLES] = currProcPoolData[poolID][UTILIZED_POOL_CYCLES];
					prevProcPoolData[poolID][POOL_TIME_CYCLES] = currProcPoolData[poolID][POOL_TIME_CYCLES];
					prevProcPoolData[poolID][TOTAL_POOL_CYCLES] = currProcPoolData[poolID][TOTAL_POOL_CYCLES];
					
					continue;
					
				} 
				
				
				if (resource_type.equals("lpar")) {
					
					
					String mode = dp.getStringValue("curr_proc_mode")[0];
					int id; 
					
					if (managerType == M_IVM)
						id = getLparIdById(dp.getStringValue("lpar_id")[0]);
					else
						id = getLparIdByName(dp.getStringValue("lpar_name")[0]);
					
					if (id<0) {
						System.out.println("Skipping unknown LPAR " + dp.getStringValue("lpar_name")[0]);
						continue;
					}
					
					// TESTING
					boolean dedicated = false;
					if (!mode.equals("shared"))
						dedicated = true;
									
					
					if (type==HOURLY) {
						if (!dedicated)
							lparEnt[id].addHourData(sample, 
								Float.parseFloat(dp.getStringValue("curr_proc_units")[0]) );
						else lparEnt[id].addHourData(sample, 
								Float.parseFloat(dp.getStringValue("curr_procs")[0]) );
						lparVP[id].addHourData(sample,
								Float.parseFloat(dp.getStringValue("curr_procs")[0]));
					} else if (type==DAILY) {
						if (!dedicated)
							lparEnt[id].addDayData(sample, 
								Float.parseFloat(dp.getStringValue("curr_proc_units")[0]) );
						else
							lparEnt[id].addDayData(sample, 
								Float.parseFloat(dp.getStringValue("curr_procs")[0]) );
						lparVP[id].addDayData(sample,
								Float.parseFloat(dp.getStringValue("curr_procs")[0]));
					}
					
					if (!dedicated) {
						s = dp.getStringValue("curr_sharing_mode")[0];
						if (s.equals("uncap"))
							cap = false;
						else
							cap = true;
					} else
							cap = true;
					
					if (!dedicated) {
						poolName = dp.getStringValue("curr_shared_proc_pool_name")[0];
						if (poolName==null)
							poolName = "DefaultPool";
					} else
						poolName = "DefaultPool";
						
				
					
					if (type==HOURLY) 
						lparStatus[id].addHourData(sample, managedSystem[num].getVarValues("name")[0], poolName, cap);
					else if (type==DAILY) 
						lparStatus[id].addDayData(sample, managedSystem[num].getVarValues("name")[0], poolName, cap);
						
					
					bi1 = new BigDecimal(dp.getStringValue("capped_cycles")[0]);
					bi2 = new BigDecimal(dp.getStringValue("uncapped_cycles")[0]);
					currLparData[id][LPAR_CYCLES] = bi1.add(bi2);
					bi1 = new BigDecimal(dp.getStringValue("time_cycles")[0]);
					currLparData[id][LPAR_TIME_CYCLES] = bi1;
					
					if ( prevLparSample[id] != null  &&
							prevLparData[id][LPAR_CYCLES]!=null && 
							prevLparData[id][LPAR_TIME_CYCLES]!=null &&
							currLparData[id][LPAR_CYCLES]!=null &&
							currLparData[id][LPAR_TIME_CYCLES]!=null &&
							prevLparData[id][LPAR_TIME_CYCLES].compareTo(currLparData[id][LPAR_TIME_CYCLES])>0 &&
							prevLparData[id][LPAR_CYCLES].compareTo(currLparData[id][LPAR_CYCLES])>=0 ) {
						
						bi1 = prevLparData[id][LPAR_CYCLES].subtract(currLparData[id][LPAR_CYCLES]);
						bi2 = prevLparData[id][LPAR_TIME_CYCLES].subtract(currLparData[id][LPAR_TIME_CYCLES]);
						if (bi2.compareTo(zero)!=0) {
							bi3=bi1.divide(bi2,new MathContext(5));
							
							if (type==HOURLY) 
								lparPC[id].addHourData(prevLparSample[id], bi3.floatValue());
							else if (type==DAILY) 
								lparPC[id].addDayData(prevLparSample[id], bi3.floatValue());
						}		
					}
					
					// Store data in previous record
					prevLparData[id][LPAR_CYCLES] = currLparData[id][LPAR_CYCLES];
					prevLparData[id][LPAR_TIME_CYCLES] = currLparData[id][LPAR_TIME_CYCLES];
					prevLparSample[id] = sample;
					
					continue;
						
				}
				 	
			}	
			
		} 
		catch (IOException ioe) {	
			System.out.println("Loader.load_single_stats: IOException");
			System.out.println(ioe);
		}
		
		System.out.print(".");
				
	}
	
	private void createExcel(String excelPath, String htmlDir, String csvDir) {
		WritableSheet sheet;
		int num;
		
		if (managedSystem==null) {
			System.out.println("No managed systems are present: aborting report generation.");
			return;
		}
		
		System.out.print("Starting Excel file creation. ");
		
		try {
			workbook = Workbook.createWorkbook(new File(excelPath));
		} catch (IOException ioe) {
			System.out.println("Error in creating Excel file "+excelPath);
			return;
		}
		
	
		
		try {
			
			num=0; 
			
			sheet = workbook.createSheet("Header", 				num++);
			//createHeader(sheet);
			createHeaderExcel(sheet);
			if (htmlDir!=null)
				createHeaderHTML(htmlDir + File.separatorChar + header_html);
			if (csvDir!=null)
				createHeaderCSV(csvDir + File.separatorChar + header_csv);
			
			if (managerType == M_HMC) {
				sheet = workbook.createSheet("HMC", 			num++);
				//createHMCSheet(sheet);
				createHMCExcel(sheet);
				if (htmlDir!=null)
					createHMCHTML(htmlDir + File.separatorChar + hmc_html);
				if (csvDir!=null)
					createHMCCSV(csvDir + File.separatorChar + hmc_csv);
			}
			
			sheet = workbook.createSheet("System_Summary", 		num++);
			if (rowMode) {
				//createSystemsSheetRowBased(sheet);
				createSystemsSheetExcel(sheet);
				if (htmlDir!=null)
					createSystemsSheetHTML(htmlDir + File.separatorChar + systems_html);
				if (csvDir!=null)
					createSystemsSheetCSV(csvDir + File.separatorChar + systems_csv);
			} else
				createSystemsSheet(sheet);
			
			sheet = workbook.createSheet("Ent_Sys_Pools", 		num++);
			createSysPoolRowBasedExcel(sheet);
			if (htmlDir!=null)
				createSysPoolRowBasedHTML(htmlDir + File.separatorChar + syspool_html);
			if (csvDir!=null)
				createSysPoolRowBasedCSV(csvDir + File.separatorChar + syspool_csv);
			 
			
			sheet = workbook.createSheet("OnOff CoD", 	num++);			
			createOnOffSheetExcel(sheet);	
			if (htmlDir!=null)
				createOnOffSheetHTML(htmlDir + File.separatorChar + onoff_html);	
			if (csvDir!=null)
				createOnOffSheetCSV(csvDir + File.separatorChar + onoff_csv);
			
			sheet = workbook.createSheet("CoD Events", 	num++);			
			createCoDLogSheetExcel(sheet);	
			if (htmlDir!=null)
				createCoDLogSheetHTML(htmlDir + File.separatorChar + codlog_html);	
			if (csvDir!=null)
				createCoDLogSheetCSV(csvDir + File.separatorChar + codlog_csv);
			
			sheet = workbook.createSheet("LPAR_Summary", 		num++);
			if (rowMode) {
				//createLparSheetRowBased(sheet);
				createLparSheetRowBasedExcel(sheet);
				if (htmlDir!=null)
					createLparSheetRowBasedHTML(htmlDir + File.separatorChar + lpar_html);
				if (csvDir!=null)
					createLparSheetRowBasedCSV(csvDir + File.separatorChar + lpar_csv);
			} else
				createLparSheet(sheet);
			
			sheet = workbook.createSheet("LPAR_Profiles", 		num++);
			//createProfileSheet(sheet);
			createProfileSheetExcel(sheet);
			if (htmlDir!=null)
				createProfileSheetHTML(htmlDir + File.separatorChar + profile_html);
			if (csvDir!=null)
				createProfileSheetCSV(csvDir + File.separatorChar + profile_csv);
			
			sheet = workbook.createSheet("LPAR_CPU", 			num++);
			if (rowMode) {
				//createCPUSheetRowBased(sheet);
				createCPUSheetRowBasedExcel(sheet);
				if (htmlDir!=null)
					createCPUSheetRowBasedHTML(htmlDir + File.separatorChar + cpu_html);
				if (csvDir!=null)
					createCPUSheetRowBasedCSV(csvDir + File.separatorChar + cpu_csv);
			} else
				createCPUSheet(sheet);
			
			sheet = workbook.createSheet("LPAR_MEM", 			num++);
			if (rowMode) {
				//createMemSheetRowBased(sheet);
				createMemSheetRowBasedExcel(sheet);
				if (htmlDir!=null)
					createMemSheetRowBasedHTML(htmlDir + File.separatorChar + mem_html);
				if (csvDir!=null)
					createMemSheetRowBasedCSV(csvDir + File.separatorChar + mem_csv);
			} else
				createMemSheet(sheet);
			
			sheet = workbook.createSheet("Physical_Slots", 		num++);			
			if (rowMode) {
				createSystemSlotSheetRowBasedExcel(sheet);	
				if (htmlDir!=null)
					createSystemSlotSheetRowBasedHTML(htmlDir + File.separatorChar + slots_html);	
				if (csvDir!=null)
					createSystemSlotSheetRowBasedCSV(csvDir + File.separatorChar + slots_csv);	
			} else
				createSystemSlotSheet(sheet);
			
			
			sheet = workbook.createSheet("IO_slot_children", 	num++);			
			//createIOChildrenSheet(sheet);	
			createIOChildrenSheetExcel(sheet);	
			if (htmlDir!=null)
				createIOChildrenSheetHTML(htmlDir + File.separatorChar + iochildren_html);	
			if (csvDir!=null)
				createIOChildrenSheetCSV(csvDir + File.separatorChar + iochildren_csv);	
					
			sheet = workbook.createSheet("Virtual_Ethernet", 	num++);
			//createVEthSheet(sheet);
			createVEthSheetExcel(sheet);
			if (htmlDir!=null)
				createVEthSheetHTML(htmlDir + File.separatorChar + veth_html);
			if (csvDir!=null)
				createVEthSheetCSV(csvDir + File.separatorChar + veth_csv);
			
			sheet = workbook.createSheet("Virtual_SCSI", 		num++);
			if (rowMode) {
				//createVSCSISheetRowBased(sheet);
				createVSCSISheetRowBasedExcel(sheet);
				if (htmlDir!=null)
					createVSCSISheetRowBasedHTML(htmlDir + File.separatorChar + vscsi_html);
				if (csvDir!=null)
					createVSCSISheetRowBasedCSV(csvDir + File.separatorChar + vscsi_csv);
			} else
				createVSCSISheet(sheet);
			
			sheet = workbook.createSheet("VSCSI_Map", 			num++);		
			if (rowMode) {
				//createVSCSIMapSheetRowBased(sheet);
				createVSCSIMapSheetRowBasedExcel(sheet);
				if (htmlDir!=null)
					createVSCSIMapSheetRowBasedHTML(htmlDir + File.separatorChar + vscsimap_html);
				if (csvDir!=null)
					createVSCSIMapSheetRowBasedCSV(csvDir + File.separatorChar + vscsimap_csv);
			} else
				createVSCSIMapSheet(sheet);
			
			sheet = workbook.createSheet("Virtual_Fibre", 		num++);
			if (rowMode) {
				//createVFCSheetRowBased(sheet);
				createVFCSheetRowBasedExcel(sheet);
				if (htmlDir!=null)
					createVFCSheetRowBasedHTML(htmlDir + File.separatorChar + vfc_html);
				if (csvDir!=null)
					createVFCSheetRowBasedCSV(htmlDir + File.separatorChar + vfc_csv);
			} else
				createVFCSheet(sheet);
			
			sheet = workbook.createSheet("VIOS disks", num++);
			//createViosDiskSheet(sheet);
			createViosDiskSheetExcel(sheet);
			if (htmlDir!=null)
				createViosDiskSheetHTML(htmlDir + File.separatorChar + viosdisks_html);
			if (csvDir!=null)
				createViosDiskSheetCSV(csvDir + File.separatorChar + viosdisks_csv);
			
			sheet = workbook.createSheet("SEA", 		num++);
			//try { 
				//createSEASheet(sheet);
				createSEASheetExcel(sheet);
				if (htmlDir!=null)
					createSEASheetHTML(htmlDir + File.separatorChar + sea_html);
				if (csvDir!=null)
					createSEASheetCSV(csvDir + File.separatorChar + sea_csv);
			//} catch (Exception e) {System.out.println("Exception in SEA sheet.");}
			
			sheet = workbook.createSheet("FC", 		num++);
			try { 
				//createFCSheet(sheet);
				createFCSheetExcel(sheet); 
				if (htmlDir!=null)
					createFCSheetHTML(htmlDir + File.separatorChar + pfc_html); 
				if (csvDir!=null)
					createFCSheetCSV(csvDir + File.separatorChar + pfc_csv); 
			} catch (Exception e) {System.out.println("Exception in FC sheet.");}
			
			
			sheet = workbook.createSheet("SW_cores", 			num++);
			createSWSheet(sheet);
			
			if (produceStatistics) {
				sheet = workbook.createSheet("CPU_Pool_Usage",		num++);
				createSysPoolUsageSheetExcel(sheet);	
				if (htmlDir!=null)
					createSysPoolUsageSheetHTML(htmlDir + File.separatorChar + poolcpu_html);
				if (csvDir!=null)
					createSysPoolUsageSheetCSV(csvDir + File.separatorChar + poolcpu_csv);
				sheet = workbook.createSheet("Sys_RAM_Usage", 			num++);
				createSysRAMUsageSheetExcel(sheet);	
				if (htmlDir!=null)
					createSysRAMUsageSheetHTML(htmlDir + File.separatorChar + sysram_html);
				if (csvDir!=null)
					createSysRAMUsageSheetCSV(csvDir + File.separatorChar + sysram_csv);
				sheet = workbook.createSheet("LPAR_CPU_Usage", 			num++);
				createLparCoreUsageSheetExcel(sheet);
				if (htmlDir!=null)
					createLparCoreUsageSheetHTML(htmlDir + File.separatorChar + lparcpu_html);
				if (csvDir!=null)
					createLparCoreUsageSheetCSV(csvDir + File.separatorChar + lparcpu_csv);
				sheet = workbook.createSheet("CPU_Pool_Daily_Usage",		num++);
				createSysPoolDailyUsageSheetExcel(sheet);
				if (htmlDir!=null)
					createSysPoolDailyUsageSheetHTML(htmlDir + File.separatorChar + pooldaily_html);
				if (csvDir!=null)
					createSysPoolDailyUsageSheetCSV(csvDir + File.separatorChar + pooldaily_csv);
							
				final byte MAXCOL = 25;
				int from,to,count;
				from=0;
				count=1;
				while (from<lparNames.length) {
					to=from+MAXCOL-1;
					if (to>=lparNames.length)
						to=lparNames.length-1;
					sheet = workbook.createSheet("LPAR_Daily_Usage_N"+count,		num++);
					createLparDailyUsageSheetExcel(sheet,from,to);
					if (htmlDir!=null)
						createLparDailyUsageSheetHTML(htmlDir + File.separatorChar + count + "_" + lpardaily_html,from,to,count);
					if (csvDir!=null)
						createLparDailyUsageSheetCSV(csvDir + File.separatorChar + count + "_" + lpardaily_csv,from,to,count);
					count++;
					from=to+1;
				}
				
				sheet = workbook.createSheet("CPU_Pool_Hourly_Usage",		num++);
				createSysPoolHourlyUsageSheetExcel(sheet);
				if (htmlDir!=null)
					createSysPoolHourlyUsageSheetHTML(htmlDir + File.separatorChar + poolhourly_html);
				if (csvDir!=null)
					createSysPoolHourlyUsageSheetCSV(csvDir + File.separatorChar + poolhourly_csv);
				
				from=0;
				count=1;
				while (from<lparNames.length) {
					to=from+MAXCOL-1;
					if (to>=lparNames.length)
						to=lparNames.length-1;
					sheet = workbook.createSheet("LPAR_Hourly_Usage_N"+count,		num++);
					createLparHourlyUsageSheetExcel(sheet,from,to);
					if (htmlDir!=null)
						createLparHourlyUsageSheetHTML(htmlDir + File.separatorChar + count + "_" + lparhourly_html,from,to,count);
					if (csvDir!=null)
						createLparHourlyUsageSheetCSV(csvDir + File.separatorChar + count + "_" + lparhourly_csv,from,to,count);
					count++;
					from=to+1;
				}
				
				if (htmlDir!=null) {
					createSystemHtmlStructure(htmlDir);
					//addButton("System Graphs",htmlDir + File.separatorChar + sysperfindex_html);
					addButton("System Graphs",sysperfindex_html);
					
					createPoolHtmlStructure(htmlDir);
					addButton("SubPool Graphs",poolperfindex_html);
					
					createLPARHtmlStructure(htmlDir);
					addButton("LPAR Graphs",lparperfindex_html);
				}
				
			}
			
			/*
			sheet = workbook.createSheet("Debug Zone", 9);
			
			sheet = workbook.createSheet("Systems", 10);
			createGlobalSystemsSheet(sheet);
			sheet = workbook.createSheet("Slots", 11);
			createSlotsSheet(sheet);
			sheet = workbook.createSheet("CPU", 12);
			createGlobalCPUSheet(sheet);
			sheet = workbook.createSheet("Memory", 13);
			createGlobalMemSheet(sheet);
			*/
		}
		catch (RowsExceededException ree) {}
		catch (WriteException we) {}
		
		try {
			workbook.write();
			workbook.close();
		} 
		catch (IOException ioe) { }
		catch (WriteException we) { }
		
		System.out.println("Done: "+excelPath);
		
		
		if (htmlDir!=null) {
			createIndexHtml(htmlDir);
			createMenuHtml(htmlDir);
		}
		
		
	}
	
	
	
	private void createCSVfiles(String dirName, String excelName, String hmcName) {
		
		File f;
		Sheet s;
		OutputStream os;
		String encoding = "UTF8";
		OutputStreamWriter osw;
		BufferedWriter bw;
		Cell row[];
		int sheet, i;
		WorkbookSettings ws;
		Workbook w;
		File dir;
		
		
		
		try {		
			//Excel document to be imported
			ws = new WorkbookSettings();
		    ws.setLocale(new Locale("en", "EN"));
		    w = Workbook.getWorkbook(new File(excelName),ws);
		    
		    // Create directory
		    dir = new File(dirName);
			if (!dir.isDirectory() && !dir.mkdir()) {
				System.out.println("Error: can not create directory "+dirName);
				System.out.println("Skipping CSV creation");
				return;
			}
			
			System.out.print("Starting CSV file creation: ");
		    
			// Gets the sheets from workbook
		    for (sheet = 0; sheet < w.getNumberOfSheets(); sheet++) {
		    	  // Get sheet
		    	  s = w.getSheet(sheet);
		    	      	  
		    	  // Create file to store CSV
		    	  f = new File(dirName + File.separatorChar + hmcName + "_" + s.getName() + ".csv");
		    	  os = (OutputStream)new FileOutputStream(f);
		    	  osw = new OutputStreamWriter(os, encoding);
		    	  bw = new BufferedWriter(osw);
		    	  
		    	  // Gets the cells from sheet
		    	  for (i = 0 ; i < s.getRows() ; i++) {
		    		  row = s.getRow(i);
		    		  
		    		  if (row.length > 0) {
		    			  bw.write(row[0].getContents());
		    			  for (int j = 1; j < row.length; j++) {
		    				  bw.write(csvSeparator);
		    	              bw.write(row[j].getContents());
		    			  }
		    		  }
		    		  bw.newLine();
		    	  }
		    	  bw.flush();
			      bw.close();
			      
			      System.out.print(".");
		    }
		    
		    System.out.println(" Done!");
		    
		} catch (IOException ioe) {
			System.out.println("Error in creating CSV files.");
		} catch (BiffException be) {
			System.out.println("Error in creating CSV files.");
		}
	}
	
	
	
	private WritableCellFormat formatLabel(int map) throws WriteException{
		WritableCellFormat wcf;
		
		if ( (map & BOLD) != 0 )
			wcf = new WritableCellFormat(new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD, true));
		else
			wcf = new WritableCellFormat(new WritableFont(WritableFont.ARIAL, 10));
		
		return setParam(wcf,map);
	}
	
	private WritableCellFormat formatInt(int map) throws WriteException {
		WritableCellFormat wcf = new WritableCellFormat (NumberFormats.INTEGER);;

		return setParam(wcf,map);
	}
	
	private WritableCellFormat formatPerc(int map) throws WriteException {
		NumberFormat cuspercent = new NumberFormat("0.00%");
		WritableCellFormat wcf = new WritableCellFormat (cuspercent);;

		return setParam(wcf,map);
	}
	
	private WritableCellFormat formatFloat(int map) throws WriteException {
		WritableCellFormat wcf;
		
		if ( (map & BOLD) != 0 )
			wcf = new WritableCellFormat(new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD, true), NumberFormats.FORMAT3);
		else
			wcf = new WritableCellFormat (NumberFormats.FORMAT3);;

		return setParam(wcf,map);
	}
	
	private WritableCellFormat setParam(WritableCellFormat wcf, int map) throws WriteException {
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
		
	
		
		return wcf;
	}
	
	
	private void createLparSheet(WritableSheet sheet)  throws RowsExceededException, WriteException {
		GenericData lpar[];
		int row;
		int i,j;
		String s[];
		Label label;
		int lparNameSize=0;
		int osNameSize=0;
				
		row = 0;
		
		for (i=0; i<managedSystem.length; i++) {
			
			/*
			 * Show start of system
			 */
			s = managedSystem[i].getVarValues("name");
			sheet.mergeCells(0, row, 8, row+1);
			addLabel(sheet,0,row,s[0],formatLabel(BOLD|CENTRE|VCENTRE|GREEN));
			row++;
			row++;
						
			/*
			 * Setup titles
			 */ 
			addLabel(sheet,0,row,"Name",formatLabel(BOLD|B_ALL_MED));
			addLabel(sheet,1,row,"ID",formatLabel(BOLD|B_ALL_MED));
			addLabel(sheet,2,row,"Status",formatLabel(BOLD|B_ALL_MED));
			addLabel(sheet,3,row,"Environment",formatLabel(BOLD|B_ALL_MED));
			addLabel(sheet,4,row,"OS Version",formatLabel(BOLD|CENTRE|B_ALL_MED));
			addLabel(sheet,5,row,"Pool data available",formatLabel(BOLD|B_ALL_MED));
			addLabel(sheet,6,row,"Proc mode",formatLabel(BOLD|B_ALL_MED));
			addLabel(sheet,7,row,"RMC IP",formatLabel(BOLD|B_ALL_MED));
			addLabel(sheet,8,row,"RMC State",formatLabel(BOLD|B_ALL_MED));
			
			row++;			
			
			/*
			 * Write variables
			 */
			lpar = managedSystem[i].getObjects(CONFIG_LPAR);
			if (lpar==null) {
				row +=2;
				continue;
			}
			
			for (j=0; j<lpar.length; j++) {
				
				s = lpar[j].getVarValues("name");
				addLabel(sheet, 0, row, s, 0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				if (s[0].length()>lparNameSize)
					lparNameSize=s[0].length();
				
				addNumber(sheet, 1, row, lpar[j].getVarValues("lpar_id"),0, formatInt(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				addLabel(sheet, 2, row, lpar[j].getVarValues("state"),0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				addLabel(sheet, 3, row, lpar[j].getVarValues("lpar_env"),0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				
				s = lpar[j].getVarValues("os_version");
				addLabel(sheet, 4, row, s,0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				if (s[0].length()>osNameSize)
					osNameSize=s[0].length();
				
				s = lpar[j].getVarValues("allow_perf_collection");
				if (s==null) {
					// Try the deprecated value
					s = lpar[j].getVarValues("shared_proc_pool_util_auth");
				}
				if (s==null)
					label = new Label(5,row,"N/A",formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				else if (s[0].equals("1"))
					label = new Label(5,row,"true",formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				else
					label = new Label(5,row,"false",formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				sheet.addCell(label);
									
				addLabel(sheet, 6, row, lpar[j].getVarValues("curr_lpar_proc_compat_mode"),0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				
				addLabel(sheet, 7, row, lpar[j].getVarValues("rmc_ipaddr"),0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				addLabel(sheet, 8, row, lpar[j].getVarValues("rmc_state"),0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
							
				row++;				
			}
			
			row +=2;
		}
		
		
		if (lparNameSize==0)
			lparNameSize=10;
		sheet.setColumnView(0, lparNameSize+2);
		sheet.setColumnView(1, 5);
		sheet.setColumnView(2, 12);
		sheet.setColumnView(3, 13);
		sheet.setColumnView(4, osNameSize);
		sheet.setColumnView(6, 18);
		sheet.setColumnView(7, 15);
		sheet.setColumnView(8, 15);
	}
	
	
	private void createLparSheetRowBasedExcel(WritableSheet sheet) {
		DataSheet ds = createLparSheetRowBased();
		if (ds!=null)
			ds.createExcelSheet(sheet);
	}
	
	private void createLparSheetRowBasedHTML(String fileName) {
		DataSheet ds = createLparSheetRowBased();
		if (ds!=null) {
			ds.createHTMLSheet(fileName);
			addButton("LPAR Summary",new File(fileName).getName());
		}
	}
	
	private void createLparSheetRowBasedCSV(String fileName) {
		DataSheet ds = createLparSheetRowBased();
		if (ds!=null) {
			ds.setSeparator(csvSeparator);
			ds.createCSVSheet(fileName);
		}
	}
	
	private DataSheet createLparSheetRowBased() {
		DataSheet ds = new DataSheet();
		GenericData lpar[];
		int col, row;
		int i,j;
		String s[];
		String sysName[];
		String serial[];
		String str;
		
		int size[]=new int[20];
		int n;
		for (i=0; i<size.length; i++)
			size[i] = 0;
				
		row = 0;
		col = 0;
		
		/*
		 * Setup titles
		 */ 
		n = ds.addLabel(col,row,"Name",BOLD|B_ALL_MED|GREEN|DIAG45); col++;
		n = ds.addLabel(col,row,"ID",BOLD|B_ALL_MED|GREEN|DIAG45); col++;
		n = ds.addLabel(col,row,"Status",BOLD|B_ALL_MED|GREEN|DIAG45); col++;
		n = ds.addLabel(col,row,"Environment",BOLD|B_ALL_MED|GREEN|DIAG45); col++;
		n = ds.addLabel(col,row,"OS Version",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45); col++;
		n = ds.addLabel(col,row,"Pool data available",BOLD|B_ALL_MED|GREEN|DIAG45); col++;
		n = ds.addLabel(col,row,"Proc mode",BOLD|B_ALL_MED|GREEN|DIAG45); col++;
		n = ds.addLabel(col,row,"RMC IP",BOLD|B_ALL_MED|GREEN|DIAG45); col++;
		n = ds.addLabel(col,row,"RMC State",BOLD|B_ALL_MED|GREEN|DIAG45); col++;
		n = ds.addLabel(col,row,"Default profile",BOLD|B_ALL_MED|GREEN|DIAG45); col++;
		n = ds.addLabel(col,row,"Current profile",BOLD|B_ALL_MED|GREEN|DIAG45); col++;
		n = ds.addLabel(col,row,"Migration disabled",BOLD|B_ALL_MED|GREEN|DIAG45); col++;
		n = ds.addLabel(col,row,"Auto start",BOLD|B_ALL_MED|GREEN|DIAG45); col++;
		n = ds.addLabel(col,row,"Suspend capable",BOLD|B_ALL_MED|GREEN|DIAG45); col++;
		n = ds.addLabel(col,row,"Remote restart capable",BOLD|B_ALL_MED|GREEN|DIAG45); col++;
		n = ds.addLabel(col,row,"Simplified remote restart capable",BOLD|B_ALL_MED|GREEN|DIAG45); col++;
		n = ds.addLabel(col,row,"Remote restart status",BOLD|B_ALL_MED|GREEN|DIAG45); col++;
		n = ds.addLabel(col,row,"Sync current profile",BOLD|B_ALL_MED|GREEN|DIAG45); col++;
		n = ds.addLabel(col,row,"Managed System Name",BOLD|B_ALL_MED|GREEN|DIAG45); col++;
		n = ds.addLabel(col,row,"Managed System Serial",BOLD|B_ALL_MED|GREEN|DIAG45); col++;
		
		row++;	
		
		for (i=0; i<managedSystem.length; i++) {
			
			/*
			 * Show start of system
			 */
			sysName = managedSystem[i].getVarValues("name");
			serial = managedSystem[i].getVarValues("serial_num");
		
			
			/*
			 * Write variables
			 */
			lpar = managedSystem[i].getObjects(CONFIG_LPAR);
			if (lpar==null) 
				continue;
			
			
			for (j=0; j<lpar.length; j++) {
				
				col = 0;
				
				s = lpar[j].getVarValues("name");
				n = ds.addLabel(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				ds.addInteger(col, row, lpar[j].getVarValues("lpar_id"),0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); n=4; if (n>size[col]) size[col]=n; col++;
				n = ds.addLabel(col, row, lpar[j].getVarValues("state"),0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				n = ds.addLabel(col, row, lpar[j].getVarValues("lpar_env"),0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				s = lpar[j].getVarValues("os_version");
				n = ds.addLabel(col, row, s,0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				s = lpar[j].getVarValues("allow_perf_collection");
				if (s==null) {
					// Try the deprecated value
					s = lpar[j].getVarValues("shared_proc_pool_util_auth");
				}
				if (s==null)
					str = "N/A";
				else if (s[0].equals("1"))
					str = "true";
				else
					str = "false";
				n = ds.addLabel(col, row, str, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
									
				n = ds.addLabel(col, row, lpar[j].getVarValues("curr_lpar_proc_compat_mode"),0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				n = ds.addLabel(col, row, lpar[j].getVarValues("rmc_ipaddr"),0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				n = ds.addLabel(col, row, lpar[j].getVarValues("rmc_state"),0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				n = ds.addLabel(col, row, lpar[j].getVarValues("default_profile"),0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				n = ds.addLabel(col, row, lpar[j].getVarValues("curr_profile"),0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				s = lpar[j].getVarValues("migration_disabled");
				if (s==null || s[0].equals("0") )
					str = "false";
				else 
					str = "true";				
				n = ds.addLabel(col, row, str, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				s = lpar[j].getVarValues("auto_start");
				if (s==null || s[0].equals("0") )
					str = "false";
				else 
					str = "true";				
				n = ds.addLabel(col, row, str, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				s = lpar[j].getVarValues("suspend_capable");
				if (s==null || s[0].equals("0") )
					str = "false";
				else 
					str = "true";				
				n = ds.addLabel(col, row, str, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				s = lpar[j].getVarValues("remote_restart_capable");
				if (s==null || s[0].equals("0") )
					str = "false";
				else 
					str = "true";				
				n = ds.addLabel(col, row, str, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				s = lpar[j].getVarValues("simplified_remote_restart_capable");
				if (s==null || s[0].equals("0") )
					str = "false";
				else 
					str = "true";				
				n = ds.addLabel(col, row, str, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				s = lpar[j].getVarValues("remote_restart_status");
				if (s==null)
					str = "N/A";
				else 
					str = s[0];		
				n = ds.addLabel(col, row, str, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				s = lpar[j].getVarValues("sync_current_profile");
				if (s==null || s[0].equals("0") )
					str = "false";
				else 
					str = "true";		
				n = ds.addLabel(col, row, str, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				n = ds.addLabel(col, row, sysName,0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				n = ds.addLabel(col, row, serial,0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
							
				row++;				
			}
		}
		
		for (i=0; i<size.length; i++)
			ds.setColSize(i, size[i]+2);

		return ds;
	}
	

	private void createLparSheetRowBased(WritableSheet sheet)  throws RowsExceededException, WriteException {
		GenericData lpar[];
		int row;
		int i,j;
		String s[];
		Label label;
		int lparNameSize=0;
		int osNameSize=0;
		String sysName[];
		String serial[];
				
		row = 0;
		
		/*
		 * Setup titles
		 */ 
		addLabel(sheet,0,row,"Name",formatLabel(BOLD|B_ALL_MED|GREEN));
		addLabel(sheet,1,row,"ID",formatLabel(BOLD|B_ALL_MED|GREEN));
		addLabel(sheet,2,row,"Status",formatLabel(BOLD|B_ALL_MED|GREEN));
		addLabel(sheet,3,row,"Environment",formatLabel(BOLD|B_ALL_MED|GREEN));
		addLabel(sheet,4,row,"OS Version",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN));
		addLabel(sheet,5,row,"Pool data available",formatLabel(BOLD|B_ALL_MED|GREEN));
		addLabel(sheet,6,row,"Proc mode",formatLabel(BOLD|B_ALL_MED|GREEN));
		addLabel(sheet,7,row,"RMC IP",formatLabel(BOLD|B_ALL_MED|GREEN));
		addLabel(sheet,8,row,"RMC State",formatLabel(BOLD|B_ALL_MED|GREEN));
		addLabel(sheet,9,row,"Default profile",formatLabel(BOLD|B_ALL_MED|GREEN));
		addLabel(sheet,10,row,"Current profile",formatLabel(BOLD|B_ALL_MED|GREEN));
		addLabel(sheet,11,row,"Managed System Name",formatLabel(BOLD|B_ALL_MED|GREEN));
		addLabel(sheet,12,row,"Managed System Serial",formatLabel(BOLD|B_ALL_MED|GREEN));
		
		row++;	
		
		for (i=0; i<managedSystem.length; i++) {
			
			/*
			 * Show start of system
			 */
			sysName = managedSystem[i].getVarValues("name");
			serial = managedSystem[i].getVarValues("serial_num");
		
			
			/*
			 * Write variables
			 */
			lpar = managedSystem[i].getObjects(CONFIG_LPAR);
			if (lpar==null) 
				continue;
			
			
			for (j=0; j<lpar.length; j++) {
				
				s = lpar[j].getVarValues("name");
				addLabel(sheet, 0, row, s, 0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				if (s[0].length()>lparNameSize)
					lparNameSize=s[0].length();
				
				addNumber(sheet, 1, row, lpar[j].getVarValues("lpar_id"),0, formatInt(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				addLabel(sheet, 2, row, lpar[j].getVarValues("state"),0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				addLabel(sheet, 3, row, lpar[j].getVarValues("lpar_env"),0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				
				s = lpar[j].getVarValues("os_version");
				addLabel(sheet, 4, row, s,0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				if (s[0].length()>osNameSize)
					osNameSize=s[0].length();
				
				s = lpar[j].getVarValues("allow_perf_collection");
				if (s==null) {
					// Try the deprecated value
					s = lpar[j].getVarValues("shared_proc_pool_util_auth");
				}
				if (s==null)
					label = new Label(5,row,"N/A",formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				else if (s[0].equals("1"))
					label = new Label(5,row,"true",formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				else
					label = new Label(5,row,"false",formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				sheet.addCell(label);
									
				addLabel(sheet, 6, row, lpar[j].getVarValues("curr_lpar_proc_compat_mode"),0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				
				addLabel(sheet, 7, row, lpar[j].getVarValues("rmc_ipaddr"),0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				addLabel(sheet, 8, row, lpar[j].getVarValues("rmc_state"),0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				
				addLabel(sheet, 9, row, lpar[j].getVarValues("default_profile"),0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				addLabel(sheet, 10, row, lpar[j].getVarValues("curr_profile"),0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				
				addLabel(sheet, 11, row, sysName,0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				addLabel(sheet, 12, row, serial,0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
							
				row++;				
			}
		}
		
		
		if (lparNameSize==0)
			lparNameSize=10;
		sheet.setColumnView(0, lparNameSize+2);
		sheet.setColumnView(1, 5);
		sheet.setColumnView(2, 12);
		sheet.setColumnView(3, 13);
		sheet.setColumnView(4, osNameSize);
		sheet.setColumnView(6, 18);
		sheet.setColumnView(7, 15);
		sheet.setColumnView(8, 15);
		
		sheet.setColumnView(9, 26);
		sheet.setColumnView(10, 26);
	}
	

	private void createMemSheet(WritableSheet sheet)  throws RowsExceededException, WriteException {
		GenericData lpar[], pool[];
		int row;
		int i,j;
		String s[];
		Label label;
		boolean dedicated;		
		int lparNameSize=0;
		int viosNameSize=0;
			
		row = 0;	

		for (i=0; i<managedSystem.length; i++) {
			
			/*
			 * Show start of system
			 */
			s = managedSystem[i].getVarValues("name");
			sheet.mergeCells(0, row, 12, row+1);
			addLabel(sheet,0,row,s[0],formatLabel(BOLD|CENTRE|VCENTRE|GREEN));
			row++;
			row++;
			
			/*
			 * Setup titles
			 */ 
			sheet.mergeCells(0, row, 0, row+1);
			addLabel(sheet,0,row,"Name",formatLabel(BOLD|VCENTRE|B_ALL_MED));
			
			sheet.mergeCells(1, row, 1, row+1);
			addLabel(sheet,1,row,"Mode",formatLabel(BOLD|VCENTRE|B_ALL_MED));
			
			sheet.mergeCells(2, row, 5, row);
			addLabel(sheet,2,row,"Memory (MB)",formatLabel(BOLD|B_ALL_MED));
			
			addLabel(sheet,2,row+1,"Min",formatLabel(BOLD|RIGHT|B_ALL_MED));
			addLabel(sheet,3,row+1,"Curr",formatLabel(BOLD|RIGHT|B_ALL_MED));
			addLabel(sheet,4,row+1,"Max",formatLabel(BOLD|RIGHT|B_ALL_MED));
			addLabel(sheet,5,row+1,"ExpFact",formatLabel(BOLD|RIGHT|B_ALL_MED));
			
			sheet.mergeCells(6, row, 12, row);
			addLabel(sheet,6,row,"Active Memory Sharing",formatLabel(BOLD|B_ALL_MED));
			
			addLabel(sheet,6,row+1,"Weight",formatLabel(BOLD|RIGHT|B_ALL_MED));
			addLabel(sheet,7,row+1,"Prim VIOS",formatLabel(BOLD|RIGHT|B_ALL_MED));
			addLabel(sheet,8,row+1,"Sec VIOS",formatLabel(BOLD|RIGHT|B_ALL_MED));
			addLabel(sheet,9,row+1,"Curr VIOS",formatLabel(BOLD|RIGHT|B_ALL_MED));
			addLabel(sheet,10,row+1,"Active",formatLabel(BOLD|RIGHT|B_ALL_MED));
			addLabel(sheet,11,row+1,"Max",formatLabel(BOLD|RIGHT|B_ALL_MED));
			addLabel(sheet,12,row+1,"Firmw",formatLabel(BOLD|RIGHT|B_ALL_MED));
			
			row +=2;
			
			
			/*
			 * Write variables
			 */
			lpar = managedSystem[i].getObjects(MEM_LPAR);
			if (lpar==null) {
				row +=2;
				continue;
			}
			
			for (j=0; j<lpar.length; j++) {
				
				s = lpar[j].getVarValues("lpar_name");
				addLabel(sheet, 0, row, s, 0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				if (s[0].length()>lparNameSize)
					lparNameSize=s[0].length();
					
				dedicated = true;
				s = lpar[j].getVarValues("mem_mode");
				if (s==null) {
					s = new String[1];
					s[0] = "ded";
					dedicated = true;
				} else if (s[0].equals("shared"))
					dedicated = false;
				addLabel(sheet, 1, row, s,0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				addNumber(sheet, 2, row, lpar[j].getVarValues("curr_min_mem"),0, formatInt(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				addNumber(sheet, 3, row, lpar[j].getVarValues("curr_mem"),0, formatInt(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				addNumber(sheet, 4, row, lpar[j].getVarValues("curr_max_mem"),0, formatInt(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				
				s = lpar[j].getVarValues("curr_mem_expansion");
				if (s==null) {
					sheet.addCell(new Number(5,row,0.0,formatFloat(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW)));
				} else
					addNumber(sheet, 5, row, s, 0, formatFloat(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				
				addNumber(sheet, 6, row, lpar[j].getVarValues("run_mem_weight"),0, formatInt(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				
				s = lpar[j].getVarValues("primary_paging_vios_name");
				addLabel(sheet, 7, row, s,0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				if (s!=null && s[0]!=null)
					if (s[0].length()>viosNameSize)
						viosNameSize=s[0].length();
				
				s = lpar[j].getVarValues("secondary_paging_vios_name");
				addLabel(sheet, 8, row, s,0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				if (s!=null && s[0]!=null)
					if (s[0].length()>viosNameSize)
						viosNameSize=s[0].length();
				
				s = lpar[j].getVarValues("curr_paging_vios_name");
				addLabel(sheet, 9, row, s,0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				if (s!=null && s[0]!=null)
					if (s[0].length()>viosNameSize)
						viosNameSize=s[0].length();
								
				pool = managedSystem[i].getObjects(MEM_POOL);
				
				if (!dedicated && pool!=null && pool[0]!=null) {								
					addNumber(sheet, 10, row, pool[0].getVarValues("curr_pool_mem"),0, formatInt(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					addNumber(sheet, 11, row, pool[0].getVarValues("curr_max_pool_mem"),0, formatInt(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					addNumber(sheet, 12, row, pool[0].getVarValues("sys_firmware_pool_mem"),0, formatInt(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				}
							
				row++;				
			}
			
			row +=2;
		}
		
		if (viosNameSize < 10)
			viosNameSize=8;

		
		sheet.setColumnView(0, lparNameSize+2);
		sheet.setColumnView(2, 8);
		sheet.setColumnView(3, 8);
		sheet.setColumnView(4, 8);
		sheet.setColumnView(5, 8);
		sheet.setColumnView(6, 8);
		sheet.setColumnView(7, viosNameSize+2);
		sheet.setColumnView(8, viosNameSize+2);
		sheet.setColumnView(9, viosNameSize+2);
		sheet.setColumnView(10, 8);
		sheet.setColumnView(11, 8);
		sheet.setColumnView(12, 8);	
	}
	
	
	
	

	private DataSheet createMemSheetRowBased() {
		DataSheet ds = new DataSheet();
		GenericData lpar[], pool[];
		int row, col;
		int i,j;
		String s[];
		boolean dedicated;		
		double d;	
		int size[]=new int[17];
		int n;
		
		String currentProfileName = null;
		GenericData currentProfileData = null;
		int map;
		double active, desired;
		boolean poweredOff;
			
		row = 0;
		col = 0;
		for (i=0; i<size.length; i++)
			size[i] = 0;
		
		
		/*
		 * Setup titles
		 */ 

		n = ds.addLabel(col,row,"Name",BOLD|VCENTRE|B_ALL_MED|GREEN);	if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Mode",BOLD|VCENTRE|B_ALL_MED|GREEN);	if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Min GB",BOLD|RIGHT|B_ALL_MED|GREEN);	if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Curr GB",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Max GB",BOLD|RIGHT|B_ALL_MED|GREEN);	if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"ExpFact",BOLD|RIGHT|B_ALL_MED|GREEN);	if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"AMS->",BOLD|B_ALL_MED|GREEN);	if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Weight",BOLD|RIGHT|B_ALL_MED|GREEN);	if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Prim VIOS",BOLD|RIGHT|B_ALL_MED|GREEN);	if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Sec VIOS",BOLD|RIGHT|B_ALL_MED|GREEN);	if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Curr VIOS",BOLD|RIGHT|B_ALL_MED|GREEN);	if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Active",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Max",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Firmw",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"<-AMS",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n=ds.addLabel(col,row,"Managed System Name",BOLD|CENTRE|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n=ds.addLabel(col,row,"Managed System Serial",BOLD|CENTRE|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		
		row++;
		
		

		for (i=0; i<managedSystem.length; i++) {
			/*
			 * Write variables
			 */
			lpar = managedSystem[i].getObjects(MEM_LPAR);
			if (lpar==null) 
				continue;
			
			map = B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW;
			
			for (j=0; j<lpar.length; j++) {
				
				col = 0;
				
				s = lpar[j].getVarValues("lpar_name");
				n = ds.addLabel(col, row, s, 0, map); if (n>size[col]) size[col]=n; col++;
				
				currentProfileName = getActiveProfileName(i,s[0]);
				currentProfileData = getProfileData(i,s[0], currentProfileName);
				
				poweredOff=false;
					
				dedicated = true;
				s = lpar[j].getVarValues("mem_mode");
				if (s==null) {
					s = new String[1];
					s[0] = "ded";
					dedicated = true;
				} else if (s[0].equals("shared"))
					dedicated = false;
				n = ds.addLabel(col, row, s,0, map); if (n>size[col]) size[col]=n; col++;
				
				ds.addFloatDiv1024(col, row, lpar[j].getVarValues("curr_min_mem"), 0, map); n=8; if (n>size[col]) size[col]=n; col++;
				
				if (currentProfileData!=null && !poweredOff) {
					active = Double.parseDouble(lpar[j].getVarValues("curr_mem")[0]);
					s = currentProfileData.getVarValues("desired_mem");
					if (s!=null) {
						desired = Double.parseDouble(s[0]);
						if (active!=desired)
							map = map | YELLOW;
					}
				}
				ds.addFloatDiv1024(col, row, lpar[j].getVarValues("curr_mem"), 0, map); n=8; if (n>size[col]) size[col]=n; col++;
				map = B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW;
				
				ds.addFloatDiv1024(col, row, lpar[j].getVarValues("curr_max_mem"), 0, map); n=8; if (n>size[col]) size[col]=n; col++;
			
				
				s = lpar[j].getVarValues("curr_mem_expansion");
				if (s!=null) {
					d = Double.parseDouble(s[0]);
					if (currentProfileData!=null && !poweredOff) {
						active = d;
						s = currentProfileData.getVarValues("mem_expansion");
						if (s!=null && !s[0].equals("null")) {
							desired = Double.parseDouble(s[0]);
							if (active!=desired)
								map = map | YELLOW;
						}
					}
				}
				ds.addFloat(col, row, s, 0, map); n=6; if (n>size[col]) size[col]=n; col++;
				map = B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW;

				
				// Separator
				n = ds.addLabel(col, row, " ", map|BLACK); if (n>size[col]) size[col]=n; col++;
				
				ds.addInteger(col, row, lpar[j].getVarValues("run_mem_weight"),0, map); n=8; if (n>size[col]) size[col]=n; col++;
				
				n = ds.addLabel(col, row, lpar[j].getVarValues("primary_paging_vios_name"),0, map); if (n>size[col]) size[col]=n; col++;			
				n = ds.addLabel(col, row, lpar[j].getVarValues("secondary_paging_vios_name"),0, map); if (n>size[col]) size[col]=n; col++;
				n = ds.addLabel(col, row, lpar[j].getVarValues("curr_paging_vios_name"),0, map); if (n>size[col]) size[col]=n; col++;
								
				pool = managedSystem[i].getObjects(MEM_POOL);			
				if (!dedicated && pool!=null && pool[0]!=null) {								
					ds.addInteger(col, row, pool[0].getVarValues("curr_pool_mem"),0, map);n=8; if (n>size[col]) size[col]=n; col++;
					ds.addInteger(col, row, pool[0].getVarValues("curr_max_pool_mem"),0, map);n=8; if (n>size[col]) size[col]=n; col++;
					ds.addInteger(col, row, pool[0].getVarValues("sys_firmware_pool_mem"),0, map);n=8; if (n>size[col]) size[col]=n; col++;
				} else
					col = col + 3;
				
				// Separator
				n = ds.addLabel(col, row, " ", map|BLACK);if (n>size[col]) size[col]=n; col++;
				
				n = ds.addLabel(col, row, managedSystem[i].getVarValues("name"),0, map); if (n>size[col]) size[col]=n; col++;
				n = ds.addLabel(col, row, managedSystem[i].getVarValues("serial_num"),0,  map); if (n>size[col]) size[col]=n; col++;
							
				row++;				
			}
		}
		
		for (i=0; i<size.length; i++)
			ds.setColSize(i, size[i]+2);

		return ds;
	}
	
	
	
	private void createMemSheetRowBasedExcel(WritableSheet sheet) {
		DataSheet ds = createMemSheetRowBased();
		if (ds!=null)
			ds.createExcelSheet(sheet);
	}
	
	private void createMemSheetRowBasedHTML(String fileName) {
		DataSheet ds = createMemSheetRowBased();
		if (ds!=null) {
			ds.createHTMLSheet(fileName);
			addButton("LPAR Mem",new File(fileName).getName());
		}
	}
	
	private void createMemSheetRowBasedCSV(String fileName) {
		DataSheet ds = createMemSheetRowBased();
		if (ds!=null) {
			ds.setSeparator(csvSeparator);
			ds.createCSVSheet(fileName);
		}
	}
	
	
	
	
	private DataSheet createSysPoolRowBased() {
		
		if (entPool == null)
			return null;
		
		
		DataSheet ds = new DataSheet();
		int row, col;
		int i,j;
		String s[];
		int size[]=new int[14];
		int n;
		
		int map;
		
		String str;
			
		row = 0;
		col = 0;
		for (i=0; i<size.length; i++)
			size[i] = 0;
		
		
		/*
		 * Setup titles for system pools
		 */ 

		map = BOLD|B_ALL_MED|GREEN|DIAG45;
		n = ds.addLabel(col,row,"Name",map);			col++;
		n = ds.addLabel(col,row,"ID",map);				col++;
		n = ds.addLabel(col,row,"State",map);			col++;
		n = ds.addLabel(col,row,"Grace Period",map);	col++;
		n = ds.addLabel(col,row,"Master Name",map);		col++;
		n = ds.addLabel(col,row,"Master Serial",map);	col++;
		n = ds.addLabel(col,row,"Backup Serial",map);	col++;
		n = ds.addLabel(col,row,"Mobile Procs",map);	col++;
		n = ds.addLabel(col,row,"Avail Mob Procs",map);	col++;
		n = ds.addLabel(col,row,"Unreturned Mob Procs",map);	col++;
		n = ds.addLabel(col,row,"Mobile Mem",map);	col++;
		n = ds.addLabel(col,row,"Avail Mob Mem",map);	col++;
		n = ds.addLabel(col,row,"Unreturned Mob Mem",map);	col++;
		
		row++;
		
		

		for (i=0; i<entPool.length; i++) {
			/*
			 * Write variables
			 */
			
			map = B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW;
			col=0;
			
			n = ds.addLabel(col, row, entPool[i].getVarValues("name"), 0, map); if (n>size[col]) size[col]=n; col++;
			ds.addInteger(col, row, entPool[i].getVarValues("id"), 0, map); n=4; if (n>size[col]) size[col]=n; col++;
			n = ds.addLabel(col, row, entPool[i].getVarValues("state"), 0, map); if (n>size[col]) size[col]=n; col++;
			
			str = "";
			s = entPool[i].getVarValues("grace_period_days_remaining");
			if (s!=null && s[0]!=null) 
				str = str + s[0] + "D ";
			s = entPool[i].getVarValues("grace_period_hours_remaining");
			if (s!=null && s[0]!=null)
				str = str + s[0] + "H";	
			n = ds.addLabel(col, row, str, map); if (n<5) n=5; if (n>size[col]) size[col]=n;		
			col++;
			
			n = ds.addLabel(col, row, entPool[i].getVarValues("master_mc_name"), 0, map); if (n>size[col]) size[col]=n; col++;
			n = ds.addLabel(col, row, entPool[i].getVarValues("master_mc_mtms"), 0, map); if (n>size[col]) size[col]=n; col++;
			n = ds.addLabel(col, row, entPool[i].getVarValues("backup_master_mc_mtms"), 0, map); if (n>size[col]) size[col]=n; col++;
			
			ds.addInteger(col, row, entPool[i].getVarValues("mobile_procs"), 0, map); n=4; if (n>size[col]) size[col]=n; col++;
			ds.addInteger(col, row, entPool[i].getVarValues("avail_mobile_procs"), 0, map); n=4; if (n>size[col]) size[col]=n; col++;
			ds.addInteger(col, row, entPool[i].getVarValues("unreturned_mobile_procs"), 0, map); n=4; if (n>size[col]) size[col]=n; col++;
			
			ds.addFloatDiv1024(col, row, entPool[i].getVarValues("mobile_mem"), 0, map); n=8; if (n>size[col]) size[col]=n; col++;
			ds.addFloatDiv1024(col, row, entPool[i].getVarValues("avail_mobile_mem"), 0, map); n=8; if (n>size[col]) size[col]=n; col++;
			ds.addFloatDiv1024(col, row, entPool[i].getVarValues("unreturned_mobile_mem"), 0, map); n=8; if (n>size[col]) size[col]=n; col++;
							
			row++;				
			
		}
		
		
		row += 2;
		col = 0;
		
		
		/*
		 * Setup titles for system pools's systems
		 */ 
		map = BOLD|B_ALL_MED|GREEN|DIAG45;
		n = ds.addLabel(col,row,"Pool Name",map);			col++;
		n = ds.addLabel(col,row,"System",map);				col++;
		n = ds.addLabel(col,row,"Installed CPU",map);		col++;
		n = ds.addLabel(col,row,"Inactive CPU",map);		col++;
		n = ds.addLabel(col,row,"Non mobile CPU",map);		col++;
		n = ds.addLabel(col,row,"Mobile CPU",map);			col++;
		n = ds.addLabel(col,row,"Unreturned mobile CPU",map);	col++;
		n = ds.addLabel(col,row,"CPU Grace Period",map);	col++;
		n = ds.addLabel(col,row,"Installed Mem",map);		col++;
		n = ds.addLabel(col,row,"Inactive Mem",map);		col++;
		n = ds.addLabel(col,row,"Non mobile Mem",map);		col++;
		n = ds.addLabel(col,row,"Mobile Mem",map);			col++;
		n = ds.addLabel(col,row,"Unreturned mobile Mem",map);	col++;
		n = ds.addLabel(col,row,"Mem Grace Period",map);	col++;

		row++;
		
		GenericData sys[];
		
		for (i=0; i<entPool.length; i++) {
					
			sys = entPool[i].getObjects(ENTPOOLSYS);
			
			for (j=0; j<sys.length; j++) {
				
				map = B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW;
				col=0;
				
				n = ds.addLabel(col, row, entPool[i].getVarValues("name"), 0, map); if (n>size[col]) size[col]=n; col++;
				n = ds.addLabel(col, row, sys[j].getVarValues("name"), 0, map); if (n>size[col]) size[col]=n; col++;

				ds.addInteger(col, row, sys[j].getVarValues("installed_procs"), 0, map); n=4; if (n>size[col]) size[col]=n; col++;
				ds.addInteger(col, row, sys[j].getVarValues("inactive_procs"), 0, map); n=4; if (n>size[col]) size[col]=n; col++;
				ds.addInteger(col, row, sys[j].getVarValues("non_mobile_procs"), 0, map); n=4; if (n>size[col]) size[col]=n; col++;
				ds.addInteger(col, row, sys[j].getVarValues("mobile_procs"), 0, map); n=4; if (n>size[col]) size[col]=n; col++;
				ds.addInteger(col, row, sys[j].getVarValues("unreturned_mobile_procs"), 0, map); n=4; if (n>size[col]) size[col]=n; col++;

				str = "";
				s = sys[j].getVarValues("proc_grace_period_days_remaining");
				if (s!=null && s[0]!=null) 
					str = str + s[0] + "D ";
				s = sys[j].getVarValues("proc_grace_period_hours_remaining");
				if (s!=null && s[0]!=null)
					str = str + s[0] + "H";	
				n = ds.addLabel(col, row, str, map); if (n<5) n=5; if (n>size[col]) size[col]=n;		
				col++;
				
				ds.addFloatDiv1024(col, row, sys[j].getVarValues("installed_mem"), 0, map); n=8; if (n>size[col]) size[col]=n; col++;
				ds.addFloatDiv1024(col, row, sys[j].getVarValues("inactive_mem"), 0, map); n=8; if (n>size[col]) size[col]=n; col++;
				ds.addFloatDiv1024(col, row, sys[j].getVarValues("non_mobile_mem"), 0, map); n=8; if (n>size[col]) size[col]=n; col++;
				ds.addFloatDiv1024(col, row, sys[j].getVarValues("mobile_mem"), 0, map); n=8; if (n>size[col]) size[col]=n; col++;
				ds.addFloatDiv1024(col, row, sys[j].getVarValues("unreturned_mobile_mem"), 0, map); n=8; if (n>size[col]) size[col]=n; col++;

				str = "";
				s = sys[j].getVarValues("mem_grace_period_days_remaining");
				if (s!=null && s[0]!=null) 
					str = str + s[0] + "D ";
				s = sys[j].getVarValues("mem_grace_period_hours_remaining");
				if (s!=null && s[0]!=null)
					str = str + s[0] + "H";	
				n = ds.addLabel(col, row, str, map); if (n<5) n=5; if (n>size[col]) size[col]=n;		
				col++;
				
				
				row++;
			}
			
		}
		
		
		for (i=0; i<size.length; i++)
			ds.setColSize(i, size[i]+2);

		return ds;
	}


	private void createSysPoolRowBasedExcel(WritableSheet sheet) {
		DataSheet ds = createSysPoolRowBased();
		if (ds!=null)
			ds.createExcelSheet(sheet);
	}
	
	private void createSysPoolRowBasedHTML(String fileName) {
		DataSheet ds = createSysPoolRowBased();
		if (ds!=null) {
			ds.createHTMLSheet(fileName);
			addButton("LPAR Mem",new File(fileName).getName());
		}
	}
	
	private void createSysPoolRowBasedCSV(String fileName) {
		DataSheet ds = createSysPoolRowBased();
		if (ds!=null) {
			ds.setSeparator(csvSeparator);
			ds.createCSVSheet(fileName);
		}
	}
	
	
	
	
	
	private void createMemSheetRowBased(WritableSheet sheet)  throws RowsExceededException, WriteException {
		GenericData lpar[], pool[];
		int row;
		int i,j;
		String s[];
		boolean dedicated;		
		double d;	
		int size[]=new int[17];
		int n;
		
		String currentProfileName = null;
		GenericData currentProfileData = null;
		int map;
		double active, desired;
		boolean poweredOff;
			
		row = 0;
		for (i=0; i<size.length; i++)
			size[i] = 0;
		
		
		/*
		 * Setup titles
		 */ 

		n = addLabel(sheet,0,row,"Name",formatLabel(BOLD|VCENTRE|B_ALL_MED|GREEN));
		if (n>size[0]) size[0]=n;
		n = addLabel(sheet,1,row,"Mode",formatLabel(BOLD|VCENTRE|B_ALL_MED|GREEN));
		if (n>size[1]) size[1]=n;if (n>size[0]) size[0]=n;
		n = addLabel(sheet,2,row,"Min GB",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN));
		if (n>size[2]) size[2]=n;
		n = addLabel(sheet,3,row,"Curr GB",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN));
		if (n>size[3]) size[3]=n;
		n = addLabel(sheet,4,row,"Max GB",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN));
		if (n>size[4]) size[4]=n;
		n = addLabel(sheet,5,row,"ExpFact",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN));
		if (n>size[5]) size[5]=n;
		n = addLabel(sheet,6,row,"AMS->",formatLabel(BOLD|B_ALL_MED|GREEN));
		if (n>size[6]) size[6]=n;
		n = addLabel(sheet,7,row,"Weight",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN));
		if (n>size[7]) size[7]=n;
		n = addLabel(sheet,8,row,"Prim VIOS",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN));
		if (n>size[8]) size[8]=n;
		n = addLabel(sheet,9,row,"Sec VIOS",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN));
		if (n>size[9]) size[9]=n;
		n = addLabel(sheet,10,row,"Curr VIOS",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN));
		if (n>size[10]) size[10]=n;
		n = addLabel(sheet,11,row,"Active",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN));
		if (n>size[11]) size[11]=n;
		n = addLabel(sheet,12,row,"Max",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN));
		if (n>size[12]) size[12]=n;
		n = addLabel(sheet,13,row,"Firmw",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN));
		if (n>size[13]) size[13]=n;
		n = addLabel(sheet,14,row,"<-AMS",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN));
		if (n>size[14]) size[14]=n;
		n=addLabel(sheet,15,row,"Managed System Name",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN));
		if (n>size[15]) size[15]=n;
		n=addLabel(sheet,16,row,"Managed System Serial",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN));
		if (n>size[16]) size[16]=n;
		
		row++;
		
		

		for (i=0; i<managedSystem.length; i++) {
			/*
			 * Write variables
			 */
			lpar = managedSystem[i].getObjects(MEM_LPAR);
			if (lpar==null) 
				continue;
			
			map = B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW;
			
			for (j=0; j<lpar.length; j++) {
				
				s = lpar[j].getVarValues("lpar_name");
				n = addLabel(sheet, 0, row, s, 0, formatLabel(map));
				if (n>size[0]) size[0]=n;
				
				currentProfileName = getActiveProfileName(i,s[0]);
				currentProfileData = getProfileData(i,s[0], currentProfileName);
				
				poweredOff=false;
					
				dedicated = true;
				s = lpar[j].getVarValues("mem_mode");
				if (s==null) {
					s = new String[1];
					s[0] = "ded";
					dedicated = true;
				} else if (s[0].equals("shared"))
					dedicated = false;
				n = addLabel(sheet, 1, row, s,0, formatLabel(map));
				if (n>size[1]) size[1]=n;
				
				d = Double.parseDouble(lpar[j].getVarValues("curr_min_mem")[0])/1024;
				addNumber(sheet, 2, row, d, formatFloat(map));
				
				if (currentProfileData!=null && !poweredOff) {
					active = Double.parseDouble(lpar[j].getVarValues("curr_mem")[0]);
					s = currentProfileData.getVarValues("desired_mem");
					if (s!=null) {
						desired = Double.parseDouble(s[0]);
						if (active!=desired)
							map = map | YELLOW;
					}
				}
				d = Double.parseDouble(lpar[j].getVarValues("curr_mem")[0])/1024;
				addNumber(sheet, 3, row, d, formatFloat(map));
				map = B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW;
				
				
				d = Double.parseDouble(lpar[j].getVarValues("curr_max_mem")[0])/1024;
				addNumber(sheet, 4, row, d, formatFloat(map));
				size[2]=size[3]=size[4]=8;
				
				
				s = lpar[j].getVarValues("curr_mem_expansion");
				if (s==null)
					d = 0.0;
				else {
					d = Double.parseDouble(s[0]);
					if (currentProfileData!=null && !poweredOff) {
						active = d;
						s = currentProfileData.getVarValues("mem_expansion");
						if (s!=null) {
							desired = Double.parseDouble(s[0]);
							if (active!=desired)
								map = map | YELLOW;
						}
					}
				}
				addNumber(sheet, 5, row, d, formatFloat(map));
				map = B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW;
				size[5]=8;
				
				/*
				if (s==null) {
					sheet.addCell(new Number(5,row,0.0,formatFloat(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW)));
				} else
					addNumber(sheet, 5, row, s, 0, formatFloat(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				*/
				
				// Separator
				addLabel(sheet, 6, row, "", formatLabel(map|BLACK));
				
				addNumber(sheet, 7, row, lpar[j].getVarValues("run_mem_weight"),0, formatInt(map));
				size[7]=8;
				
				s = lpar[j].getVarValues("primary_paging_vios_name");
				n = addLabel(sheet, 8, row, s,0, formatLabel(map));
				if (n>size[8]) size[8]=n;
				
				s = lpar[j].getVarValues("secondary_paging_vios_name");
				n = addLabel(sheet, 9, row, s,0, formatLabel(map));
				if (n>size[9]) size[9]=n;
				
				s = lpar[j].getVarValues("curr_paging_vios_name");
				n = addLabel(sheet, 10, row, s,0, formatLabel(map));
				if (n>size[10]) size[10]=n;
								
				pool = managedSystem[i].getObjects(MEM_POOL);
				
				if (!dedicated && pool!=null && pool[0]!=null) {								
					addNumber(sheet, 11, row, pool[0].getVarValues("curr_pool_mem"),0, formatInt(map));
					addNumber(sheet, 12, row, pool[0].getVarValues("curr_max_pool_mem"),0, formatInt(map));
					addNumber(sheet, 13, row, pool[0].getVarValues("sys_firmware_pool_mem"),0, formatInt(map));
				}
				
				// Separator
				addLabel(sheet, 14, row, "", formatLabel(map|BLACK));
				
				n = addLabel(sheet, 15, row, managedSystem[i].getVarValues("name"),0, formatLabel(map));
				if (n>size[15]) size[15]=n;
				n = addLabel(sheet, 16, row, managedSystem[i].getVarValues("serial_num"),0, formatLabel(map));
				if (n>size[16]) size[16]=n;
							
				row++;				
			}
		}
		
		for (i=0; i<size.length; i++)
			sheet.setColumnView(i, size[i]+2);
	}
	

	
	private void createSystemSlotSheet(WritableSheet sheet)  throws RowsExceededException, WriteException {
		GenericData slot[];
		int row;

		int i,j;
		String s[];				
		int lparNameSize=0;
		int descriptionSize=0;
		int locationSize=0;	
		
		row = 0;
		
		for (i=0; i<managedSystem.length; i++) {
			
			/*
			 * Show start of system
			 */
			s = managedSystem[i].getVarValues("name");
			sheet.mergeCells(0, row, 3, row+1);
			addLabel(sheet,0,row,s[0],formatLabel(BOLD|CENTRE|VCENTRE|GREEN));
			row++;
			row++;
			
			/*
			 * Setup titles
			 */ 
			addLabel(sheet,0,row,"LPAR",formatLabel(BOLD|B_ALL_MED));
			addLabel(sheet,1,row,"Adapter description",formatLabel(BOLD|B_ALL_MED));
			addLabel(sheet,2,row,"Location",formatLabel(BOLD|B_ALL_MED));
			addLabel(sheet,3,row,"drc_index",formatLabel(CENTRE|BOLD|B_ALL_MED));
			row ++;
				
			/*
			 * Write variables
			 */
			slot = managedSystem[i].getObjects(SLOT);
			if (slot==null) {
				row +=2;
				continue;
			}
			
			for (j=0; j<slot.length; j++) {
				
				s = slot[j].getVarValues("lpar_name");
				if (s!=null) {
					addLabel(sheet, 0, row, s, 0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					if (s[0].length()>lparNameSize)
						lparNameSize=s[0].length();
				} else
					addLabel(sheet, 0, row, null, 0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				
				s = slot[j].getVarValues("description");
				addLabel(sheet, 1, row, s, 0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				if (s[0].length()>descriptionSize)
					descriptionSize=s[0].length();
				
				s = slot[j].getVarValues("drc_name");
				addLabel(sheet, 2, row, s, 0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				if (s[0].length()>locationSize)
					locationSize=s[0].length();
				
				s = slot[j].getVarValues("drc_index");
				addLabel(sheet, 3, row, s, 0, formatLabel(CENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));

							
				row++;				
			}
			
			row +=2;
		}
		
		if (lparNameSize < 10)
			lparNameSize=8;

		
		sheet.setColumnView(0, lparNameSize+2);
		sheet.setColumnView(1, descriptionSize+2);
		sheet.setColumnView(2, locationSize+2);
		sheet.setColumnView(3, 10);		
	}
	
	
	private void createSystemSlotSheetRowBasedExcel(WritableSheet sheet) {
		DataSheet ds = createSystemSlotSheetRowBased();
		if (ds!=null)
			ds.createExcelSheet(sheet);
	}
	
	private void createSystemSlotSheetRowBasedHTML(String fileName) {
		DataSheet ds = createSystemSlotSheetRowBased();
		if (ds!=null) {
			ds.createHTMLSheet(fileName);
			addButton("Physical Slots",new File(fileName).getName());
		}
	}
	
	private void createSystemSlotSheetRowBasedCSV(String fileName) {
		DataSheet ds = createSystemSlotSheetRowBased();
		if (ds!=null) {
			ds.setSeparator(csvSeparator);
			ds.createCSVSheet(fileName);
		}
	}
	
	private DataSheet createSystemSlotSheetRowBased() {
		DataSheet ds = new DataSheet();
		
		GenericData slot[];
		int row,col;

		int i,j,k;
		String s[];	
		
		String currentProfileName = null;
		GenericData currentProfileData = null;
		String profSlots[];
		String drcindex;
		int map;
		boolean reqrd;
		
		int size[]=new int[9];
		int n;
		for (i=0; i<size.length; i++)
			size[i] = 0;
		
		row = 0;
		col = 0;
		
		/*
		 * Setup titles
		 */ 
		n = ds.addLabel(col,row,"LPAR",BOLD|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Reqrd",BOLD|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;		
		n = ds.addLabel(col,row,"Adapter description",BOLD|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Location",BOLD|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"drc_index",CENTRE|BOLD|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		
		n = ds.addLabel(col,row,"Profiles desiring",CENTRE|BOLD|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Profiles requiring",CENTRE|BOLD|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		
		n = ds.addLabel(col,row,"Managed System Name",BOLD|CENTRE|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Managed System Serial",BOLD|CENTRE|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		
		row ++;
		
		for (i=0; i<managedSystem.length; i++) {

			/*
			 * Write variables
			 */
			slot = managedSystem[i].getObjects(SLOT);
			if (slot==null) {
				continue;
			}
			
			for (j=0; j<slot.length; j++) {
				
				col = 0;
				map = B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW;
				
				reqrd = false;
				s = slot[j].getVarValues("lpar_name");
				if (s!=null) {
					currentProfileName = getActiveProfileName(i,s[0]);
					currentProfileData = getProfileData(i,s[0], currentProfileName);
					if (currentProfileData==null)
						profSlots=null;
					else
						profSlots = currentProfileData.getVarValues("io_slots");
					drcindex = slot[j].getVarValues("drc_index")[0];
					map = map | YELLOW;
					if (currentProfileData!=null && currentProfileData.getVarValues("all_resources")[0].equals("1")) {
						map = map ^ YELLOW;
						reqrd=true;
					}
					for (k=0; profSlots!=null && k<profSlots.length; k++)
						if (profSlots[k].startsWith(drcindex)) {
							map = map ^ YELLOW;
							if (profSlots[k].endsWith("1"))
								reqrd=true;
							break;
						}
				}
				n = ds.addLabel(col, row, s, 0, map); if (n>size[col]) size[col]=n; col++;
				
				map = B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW;
				
				if (reqrd) {
					n = ds.addLabel(col, row, "yes", map); if (n>size[col]) size[col]=n; col++;
				} else
					col++;
				
				s = slot[j].getVarValues("description");
				n = ds.addLabel(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				s = slot[j].getVarValues("drc_name");
				n = ds.addLabel(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				s = slot[j].getVarValues("drc_index");
				n = ds.addLabel(col, row, s, 0, CENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				String profiles[][] = getProfilesUsingAdapter(i,s[0]);
				if (profiles!=null) {
					ds.addMultipleLabelsWrap(col, row, profiles[0],B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW|WRAP);
					size[col]=20; col++;
					
					ds.addMultipleLabelsWrap(col, row, profiles[1],B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW|WRAP);
					size[col]=20; col++;
				} else
					col = col + 2;
				
				n = ds.addLabel(col, row, managedSystem[i].getVarValues("name"),0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				n = ds.addLabel(col, row, managedSystem[i].getVarValues("serial_num"),0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
							
				row++;				
			}
		}
		
		for (i=0; i<size.length; i++)
			ds.setColSize(i, size[i]+2);	
		
		return ds;
	}
	
	
	private void createSystemSlotSheetRowBased(WritableSheet sheet)  throws RowsExceededException, WriteException {
		GenericData slot[];
		int row,col;

		int i,j,k;
		String s[];	
		
		String currentProfileName = null;
		GenericData currentProfileData = null;
		String profSlots[];
		String drcindex;
		int map;
		boolean reqrd;
		
		int size[]=new int[9];
		int n;
		for (i=0; i<size.length; i++)
			size[i] = 0;
		
		row = 0;
		col = 0;
		
		/*
		 * Setup titles
		 */ 
		n = addLabel(sheet,col,row,"LPAR",formatLabel(BOLD|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"Reqrd",formatLabel(BOLD|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;		
		n = addLabel(sheet,col,row,"Adapter description",formatLabel(BOLD|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"Location",formatLabel(BOLD|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"drc_index",formatLabel(CENTRE|BOLD|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		
		n = addLabel(sheet,col,row,"Profiles desiring",formatLabel(CENTRE|BOLD|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"Profiles requiring",formatLabel(CENTRE|BOLD|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		
		n = addLabel(sheet,col,row,"Managed System Name",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"Managed System Serial",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		
		row ++;
		
		for (i=0; i<managedSystem.length; i++) {

			/*
			 * Write variables
			 */
			slot = managedSystem[i].getObjects(SLOT);
			if (slot==null) {
				continue;
			}
			
			for (j=0; j<slot.length; j++) {
				
				col = 0;
				map = B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW;
				
				reqrd = false;
				s = slot[j].getVarValues("lpar_name");
				if (s!=null) {
					currentProfileName = getActiveProfileName(i,s[0]);
					currentProfileData = getProfileData(i,s[0], currentProfileName);
					profSlots = currentProfileData.getVarValues("io_slots");
					drcindex = slot[j].getVarValues("drc_index")[0];
					map = map | YELLOW;
					if (currentProfileData.getVarValues("all_resources")[0].equals("1")) {
						map = map ^ YELLOW;
						reqrd=true;
					}
					for (k=0; profSlots!=null && k<profSlots.length; k++)
						if (profSlots[k].startsWith(drcindex)) {
							map = map ^ YELLOW;
							if (profSlots[k].endsWith("1"))
								reqrd=true;
							break;
						}
				}
				n = addLabel(sheet, col, row, s, 0, formatLabel(map)); if (n>size[col]) size[col]=n; col++;
				
				map = B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW;
				
				if (reqrd) {
					n = addLabel(sheet, col, row, "yes", formatLabel(map)); if (n>size[col]) size[col]=n; col++;
				} else
					col++;
				
				s = slot[j].getVarValues("description");
				n = addLabel(sheet, col, row, s, 0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW)); if (n>size[col]) size[col]=n; col++;
				
				s = slot[j].getVarValues("drc_name");
				n = addLabel(sheet, col, row, s, 0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW)); if (n>size[col]) size[col]=n; col++;
				
				s = slot[j].getVarValues("drc_index");
				n = addLabel(sheet, col, row, s, 0, formatLabel(CENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW)); if (n>size[col]) size[col]=n; col++;
				
				String profiles[][] = getProfilesUsingAdapter(i,s[0]);
				addMultipleLabelsWrap(sheet, col, row, profiles[0],formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW|WRAP));
				size[col]=20; col++;
				
				addMultipleLabelsWrap(sheet, col, row, profiles[1],formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW|WRAP));
				size[col]=20; col++;
				
				n = addLabel(sheet, col, row, managedSystem[i].getVarValues("name"),0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW)); if (n>size[col]) size[col]=n; col++;
				n = addLabel(sheet, col, row, managedSystem[i].getVarValues("serial_num"),0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW)); if (n>size[col]) size[col]=n; col++;
							
				row++;				
			}
		}
		
		for (i=0; i<size.length; i++)
			sheet.setColumnView(i, size[i]+2);		
	}
	
	
	
	private void createVEthSheetExcel(WritableSheet sheet) {
		DataSheet ds = createVEthSheet();
		if (ds!=null)
			ds.createExcelSheet(sheet);
	}
	
	private void createVEthSheetHTML(String fileName) {
		DataSheet ds = createVEthSheet();
		if (ds!=null) {
			ds.createHTMLSheet(fileName);
			addButton("Virtual Ethernet",new File(fileName).getName());
		}
	}
	
	private void createVEthSheetCSV(String fileName) {
		DataSheet ds = createVEthSheet();
		if (ds!=null) {
			ds.setSeparator(csvSeparator);
			ds.createCSVSheet(fileName);
		}
	}
	
	
	private DataSheet createVEthSheet() {
		DataSheet ds = new DataSheet();
		GenericData object[];
		int row;
		int i,j;
		String s[];
		Label label;		
		int lparNameSize=0;
		int n;
				
		row = 0;

		
		for (i=0; i<managedSystem.length; i++) {
			
			/*
			 * Show start of system
			 */
			
			
			s = managedSystem[i].getVarValues("name");
			ds.mergeCells(0, row, 9, row+1);
			n = ds.addLabel(0,row,s[0],BOLD|CENTRE|VCENTRE|GREEN);  
			row++;
			row++;
						
			object = managedSystem[i].getObjects(VSWITCH);
			if (object != null) {
			
				/*
				 * Setup titles for virtual switches
				 */ 
				ds.mergeCells(0, row, 9, row);
				n = ds.addLabel(0,row,"Virtual Switches",BOLD|CENTRE|B_ALL_MED); 
				
				n = ds.addLabel(0,row+1,"Name",BOLD|CENTRE|B_ALL_MED);
				n = ds.addLabel(0,row+2,"VLAN ids",BOLD|CENTRE|VCENTRE|B_ALL_MED);
				
				/*
				 * Write virtual switch variables
				 */
				for (j=0; j<object.length; j++) {
					
					s = object[j].getVarValues("vswitch");
					n = ds.addLabel(1+j, row+1, s,0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
					
					s = object[j].getVarValues("vlan_ids");
					ds.addMultipleLabelsWrap(1+j, row+2, s,B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW|WRAP);			
				}	
				
				row+=4;
			}
			
						
			/*
			 * Setup titles for virtual slots
			 */ 
			ds.mergeCells(0, row, 9, row);
			n = ds.addLabel(0,row,"Virtual Ethernet Slots",BOLD|B_ALL_MED|CENTRE);
			n = ds.addLabel(0,row+1,"LPAR name",BOLD|B_ALL_MED);
			n = ds.addLabel(1,row+1,"Slot",BOLD|B_ALL_MED);
			n = ds.addLabel(2,row+1,"State",BOLD|B_ALL_MED);
			n = ds.addLabel(3,row+1,"Required",BOLD|B_ALL_MED);
			n = ds.addLabel(4,row+1,"Trunk",BOLD|B_ALL_MED);
			n = ds.addLabel(5,row+1,"Trunk Priority",BOLD|B_ALL_MED);
			n = ds.addLabel(6,row+1,"MAC Addresss",BOLD|B_ALL_MED);
			n = ds.addLabel(7,row+1,"Virtual Switch",BOLD|B_ALL_MED);
			n = ds.addLabel(8,row+1,"Port VLAN id",BOLD|B_ALL_MED);
			
			n = ds.addLabel(9,row+1,"IEEE VLAN ids",BOLD|B_ALL_MED);		
			
			row +=2;
			
			/*
			 * Write virtual slot variables values
			 */
			object = managedSystem[i].getObjects(VETH);
			if (object != null) {
				for (j=0; j<object.length; j++) {
					
					s = object[j].getVarValues("lpar_name");
					n = ds.addLabel( 0, row, s,0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
					if (s!=null && s[0]!=null && s[0].length()>lparNameSize)
						lparNameSize=s[0].length();
					
					s = object[j].getVarValues("slot_num");
					ds.addInteger(1, row, s,0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
					
					s = object[j].getVarValues("state");
					if (s!=null) {
						if (s[0].equals("1")) {
							n = ds.addLabel( 2, row, "On", B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
						} else {
							n = ds.addLabel( 2, row, "Off", B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
						}
					}
					
					s = object[j].getVarValues("is_required");
					if (s!=null) {
						if (s[0].equals("1"))
							n = ds.addLabel( 3, row, "True", B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
						else
							n = ds.addLabel( 3, row, "False", B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
					}
					
					s = object[j].getVarValues("is_trunk");
					if (s!=null) {
						if (s[0].equals("1"))
							n = ds.addLabel( 4, row, "True", B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
						else
							n = ds.addLabel( 4, row, "False", B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
					}
					
					s = object[j].getVarValues("trunk_priority");
					if (s!=null) {
						ds.addInteger( 5, row, s,0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
					}
					
					s = object[j].getVarValues("mac_addr");
					n = ds.addLabel( 6, row, s,0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
					
					s = object[j].getVarValues("vswitch");
					n = ds.addLabel( 7, row, s,0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
					
					s = object[j].getVarValues("port_vlan_id");
					ds.addInteger( 8, row, s,0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
					
					s = object[j].getVarValues("addl_vlan_ids");
					ds.addMultipleLabelsWrap(9, row, s, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW|WRAP);				
								
					row++;				
				}				
			}
			
			
			row +=2;
		}

		if (lparNameSize==0)
			lparNameSize=10;
		
		ds.setColSize(0, lparNameSize+2);
		for (i=1; i<=9; i++)
			ds.setColSize(i, 20);	
		
		return ds;
	}
	
	
	
	private void createVEthSheet(WritableSheet sheet)  throws RowsExceededException, WriteException {
		GenericData object[];
		int row;
		int i,j;
		String s[];
		Label label;		
		int lparNameSize=0;
				
		row = 0;
		
		for (i=0; i<managedSystem.length; i++) {
			
			/*
			 * Show start of system
			 */
			s = managedSystem[i].getVarValues("name");
			sheet.mergeCells(0, row, 9, row+1);
			addLabel(sheet,0,row,s[0],formatLabel(BOLD|CENTRE|VCENTRE|GREEN));
			row++;
			row++;
			
			object = managedSystem[i].getObjects(VSWITCH);
			if (object != null) {
			
				/*
				 * Setup titles for virtual switches
				 */ 
				sheet.mergeCells(0, row, 9, row);
				addLabel(sheet,0,row,"Virtual Switches",formatLabel(BOLD|CENTRE|B_ALL_MED));	
				
				addLabel(sheet,0,row+1,"Name",formatLabel(BOLD|CENTRE|B_ALL_MED));
				addLabel(sheet,0,row+2,"VLAN ids",formatLabel(BOLD|CENTRE|VCENTRE|B_ALL_MED));
				
				/*
				 * Write virtual switch variables
				 */
				for (j=0; j<object.length; j++) {
					
					s = object[j].getVarValues("vswitch");
					addLabel(sheet, 1+j, row+1, s,0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					
					s = object[j].getVarValues("vlan_ids");
					addMultipleLabelsWrap(sheet, 1+j, row+2, s,formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW|WRAP));			
				}	
				
				row+=4;
			}
			
						
			/*
			 * Setup titles for virtual slots
			 */ 
			sheet.mergeCells(0, row, 9, row);
			addLabel(sheet,0,row,"Virtual Ethernet Slots",formatLabel(BOLD|B_ALL_MED|CENTRE));
			addLabel(sheet,0,row+1,"LPAR name",formatLabel(BOLD|B_ALL_MED));
			addLabel(sheet,1,row+1,"Slot",formatLabel(BOLD|B_ALL_MED));
			addLabel(sheet,2,row+1,"State",formatLabel(BOLD|B_ALL_MED));
			addLabel(sheet,3,row+1,"Required",formatLabel(BOLD|B_ALL_MED));
			addLabel(sheet,4,row+1,"Trunk",formatLabel(BOLD|B_ALL_MED));
			addLabel(sheet,5,row+1,"Trunk Priority",formatLabel(BOLD|B_ALL_MED));
			addLabel(sheet,6,row+1,"MAC Addresss",formatLabel(BOLD|B_ALL_MED));
			addLabel(sheet,7,row+1,"Virtual Switch",formatLabel(BOLD|B_ALL_MED));
			addLabel(sheet,8,row+1,"Port VLAN id",formatLabel(BOLD|B_ALL_MED));
			
			addLabel(sheet,9,row+1,"IEEE VLAN ids",formatLabel(BOLD|B_ALL_MED));		
			
			row +=2;
			
			/*
			 * Write virtual slot variables values
			 */
			object = managedSystem[i].getObjects(VETH);
			if (object != null) {
				for (j=0; j<object.length; j++) {
					
					s = object[j].getVarValues("lpar_name");
					addLabel(sheet, 0, row, s,0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					if (s!=null && s[0]!=null && s[0].length()>lparNameSize)
						lparNameSize=s[0].length();
					
					s = object[j].getVarValues("slot_num");
					addNumber(sheet, 1, row, s,0, formatInt(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					
					s = object[j].getVarValues("state");
					if (s!=null) {
						if (s[0].equals("1"))
							label = new Label(2,row,"On",formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
						else
							label = new Label(2,row,"Off",formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
						sheet.addCell(label);
					}
					
					s = object[j].getVarValues("is_required");
					if (s!=null) {
						if (s[0].equals("1"))
							label = new Label(3,row,"True",formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
						else
							label = new Label(3,row,"False",formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
						sheet.addCell(label);
					}
					
					s = object[j].getVarValues("is_trunk");
					if (s!=null) {
						if (s[0].equals("1"))
							label = new Label(4,row,"True",formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
						else
							label = new Label(4,row,"False",formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
						sheet.addCell(label);
					}
					
					s = object[j].getVarValues("trunk_priority");
					if (s!=null) {
						addLabel(sheet, 5, row, s,0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					}
					
					s = object[j].getVarValues("mac_addr");
					addLabel(sheet, 6, row, s,0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					
					s = object[j].getVarValues("vswitch");
					addLabel(sheet, 7, row, s,0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					
					s = object[j].getVarValues("port_vlan_id");
					addNumber(sheet, 8, row, s,0, formatInt(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					
					s = object[j].getVarValues("addl_vlan_ids");
					addMultipleLabelsWrap(sheet, 9, row, s,formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW|WRAP));				
								
					row++;				
				}				
			}
			
			
			row +=2;
		}

		if (lparNameSize==0)
			lparNameSize=10;
		sheet.setColumnView(0, lparNameSize+2);
		for (i=1; i<=9; i++)
			sheet.setColumnView(i, 20);	
	}
	
	
	
	private void createVSCSISheet(WritableSheet sheet)  throws RowsExceededException, WriteException {
		GenericData object[];
		int row;
		int i,j;
		String s[];
		Label label;		
		int lparNameSize=0;
		int remoteLparNameSize=0;
		
			
		row = 0;
		
		for (i=0; i<managedSystem.length; i++) {
			
			/*
			 * Show start of system
			 */
			s = managedSystem[i].getVarValues("name");
			sheet.mergeCells(0, row, 6, row+1);
			addLabel(sheet,0,row,s[0],formatLabel(BOLD|CENTRE|VCENTRE|GREEN));
			row++;
			row++;
			
			/*
			 * Setup titles for virtual slots
			 */ 			
			addLabel(sheet,0,row,"LPAR name",formatLabel(BOLD|CENTRE|B_ALL_MED));	
			addLabel(sheet,1,row,"Slot",formatLabel(BOLD|CENTRE|B_ALL_MED));
			addLabel(sheet,2,row,"State",formatLabel(BOLD|CENTRE|B_ALL_MED));	
			addLabel(sheet,3,row,"Required",formatLabel(BOLD|CENTRE|B_ALL_MED));	
			addLabel(sheet,4,row,"Type",formatLabel(BOLD|CENTRE|B_ALL_MED));			
			addLabel(sheet,5,row,"Remote LPAR",formatLabel(BOLD|CENTRE|B_ALL_MED));		
			addLabel(sheet,6,row,"Remote Slot",formatLabel(BOLD|CENTRE|B_ALL_MED));				
			
			row ++;
			
			/*
			 * Write virtual slot variables values
			 */
			object = managedSystem[i].getObjects(VSCSI);
			if (object != null) {
				for (j=0; j<object.length; j++) {
					
					s = object[j].getVarValues("lpar_name");
					addLabel(sheet, 0, row, s,0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					if (s!=null && s[0]!=null && s[0].length()>lparNameSize)
						lparNameSize=s[0].length();
					
					s = object[j].getVarValues("slot_num");
					addNumber(sheet, 1, row, s,0, formatInt(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					
					s = object[j].getVarValues("state");
					if (s!=null) {
						if (s[0].equals("1"))
							label = new Label(2,row,"On",formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
						else
							label = new Label(2,row,"Off",formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
						sheet.addCell(label);
					}
					
					s = object[j].getVarValues("is_required");
					if (s!=null) {
						if (s[0].equals("1"))
							label = new Label(3,row,"True",formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
						else
							label = new Label(3,row,"False",formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
						sheet.addCell(label);
					}
					
					s = object[j].getVarValues("adapter_type");
					addLabel(sheet, 4, row, s,0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					
					s = object[j].getVarValues("remote_lpar_name");
					addLabel(sheet, 5, row, s,0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					if (s!=null && s[0]!=null && s[0].length()>remoteLparNameSize)
						remoteLparNameSize=s[0].length();
					
					s = object[j].getVarValues("remote_slot_num");
					if (s!=null && s[0].equals("any"))
						addLabel(sheet, 6, row, s,0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					else
						addNumber(sheet, 6, row, s,0, formatInt(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
												
					row++;				
				}				
			}
			
			
			row +=2;
		}

		if (lparNameSize==0)
			lparNameSize=10;
		if (remoteLparNameSize==0)
			remoteLparNameSize=10;
		sheet.setColumnView(0, lparNameSize+2);
		sheet.setColumnView(1, 5);
		sheet.setColumnView(2, 9);
		sheet.setColumnView(3, 9);
		sheet.setColumnView(4, 9);
		sheet.setColumnView(5, remoteLparNameSize+2);
		sheet.setColumnView(6, 15);
		sheet.setColumnView(7, 125);
		
	}
	
	
	private void createVSCSISheetRowBased(WritableSheet sheet)  throws RowsExceededException, WriteException {
		GenericData object[];
		int row;
		int i,j;
		String s[];
		Label label;		
		
		int size[]=new int[9];
		int n;
		for (i=0; i<size.length; i++)
			size[i] = 0;
			
		row = 0;
		
		/*
		 * Setup titles for virtual slots
		 */ 			
		n = addLabel(sheet,0,row,"LPAR name",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN));	
		if (n>size[0]) size[0]=n;
		n = addLabel(sheet,1,row,"Slot",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN));
		if (n>size[1]) size[1]=n;
		n = addLabel(sheet,2,row,"State",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN));	
		if (n>size[2]) size[2]=n;
		n = addLabel(sheet,3,row,"Required",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN));
		if (n>size[3]) size[3]=n;
		n = addLabel(sheet,4,row,"Type",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN));	
		if (n>size[4]) size[4]=n;
		n = addLabel(sheet,5,row,"Remote LPAR",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN));	
		if (n>size[5]) size[5]=n;
		n = addLabel(sheet,6,row,"Remote Slot",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN));	
		if (n>size[6]) size[6]=n;
		n=addLabel(sheet,7,row,"Managed System Name",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN));
		if (n>size[7]) size[7]=n;
		n=addLabel(sheet,8,row,"Managed System Serial",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN));
		if (n>size[8]) size[8]=n;
		
		row++;
		
	
		for (i=0; i<managedSystem.length; i++) {
			
			/*
			 * Write virtual slot variables values
			 */
			object = managedSystem[i].getObjects(VSCSI);
			if (object != null) {
				for (j=0; j<object.length; j++) {
					
					s = object[j].getVarValues("lpar_name");
					n = addLabel(sheet, 0, row, s,0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					if (n>size[0]) size[0]=n;
					
					s = object[j].getVarValues("slot_num");
					addNumber(sheet, 1, row, s,0, formatInt(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					
					s = object[j].getVarValues("state");
					if (s!=null) {
						if (s[0].equals("1"))
							label = new Label(2,row,"On",formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
						else
							label = new Label(2,row,"Off",formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
						sheet.addCell(label);
					}
					
					s = object[j].getVarValues("is_required");
					if (s!=null) {
						if (s[0].equals("1"))
							label = new Label(3,row,"True",formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
						else
							label = new Label(3,row,"False",formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
						sheet.addCell(label);
					}
					
					s = object[j].getVarValues("adapter_type");
					n = addLabel(sheet, 4, row, s,0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					if (n>size[4]) size[4]=n;
					
					s = object[j].getVarValues("remote_lpar_name");
					n = addLabel(sheet, 5, row, s,0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					if (n>size[5]) size[5]=n;
					
					s = object[j].getVarValues("remote_slot_num");
					if (s!=null && s[0].equals("any"))
						addLabel(sheet, 6, row, s,0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					else
						addNumber(sheet, 6, row, s,0, formatInt(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					
					n = addLabel(sheet, 7, row, managedSystem[i].getVarValues("name"),0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					if (n>size[7]) size[7]=n;
					n = addLabel(sheet, 8, row, managedSystem[i].getVarValues("serial_num"),0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					if (n>size[8]) size[8]=n;
												
					row++;				
				}				
			}
		}
		
		for (i=0; i<size.length; i++)
			sheet.setColumnView(i, size[i]+2);
	}
	
	
	private void createVSCSISheetRowBasedExcel(WritableSheet sheet) {
		DataSheet ds = createVSCSISheetRowBased();
		if (ds!=null)
			ds.createExcelSheet(sheet);
	}
	
	private void createVSCSISheetRowBasedHTML(String fileName) {
		DataSheet ds = createVSCSISheetRowBased();
		if (ds!=null) {
			ds.createHTMLSheet(fileName);
			//addButton("LPAR Profiles",profile_html);
			addButton("Virtual SCSI",new File(fileName).getName());
		}
	}
	
	private void createVSCSISheetRowBasedCSV(String fileName) {
		DataSheet ds = createVSCSISheetRowBased();
		if (ds!=null) {
			ds.setSeparator(csvSeparator);
			ds.createCSVSheet(fileName);
		}
	}
	
	private DataSheet createVSCSISheetRowBased() {
		DataSheet ds = new DataSheet();
		GenericData object[];
		int row, col;
		int i,j;
		String s[];
		Label label;		
		
		int size[]=new int[9];
		int n;
		for (i=0; i<size.length; i++)
			size[i] = 0;
			
		row = 0;
		col = 0;
		
		/*
		 * Setup titles for virtual slots
		 */ 			
		n = ds.addLabel(col,row,"LPAR name",BOLD|CENTRE|B_ALL_MED|GREEN);	
		if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Slot",BOLD|CENTRE|B_ALL_MED|GREEN);
		if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"State",BOLD|CENTRE|B_ALL_MED|GREEN);	
		if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Required",BOLD|CENTRE|B_ALL_MED|GREEN);
		if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Type",BOLD|CENTRE|B_ALL_MED|GREEN);	
		if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Remote LPAR",BOLD|CENTRE|B_ALL_MED|GREEN);	
		if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Remote Slot",BOLD|CENTRE|B_ALL_MED|GREEN);	
		if (n>size[col]) size[col]=n; col++;
		n=ds.addLabel(col,row,"Managed System Name",BOLD|CENTRE|B_ALL_MED|GREEN);
		if (n>size[col]) size[col]=n; col++;
		n=ds.addLabel(col,row,"Managed System Serial",BOLD|CENTRE|B_ALL_MED|GREEN);
		if (n>size[col]) size[col]=n; col++;
		
		row++;
		
	
		for (i=0; i<managedSystem.length; i++) {
			
			/*
			 * Write virtual slot variables values
			 */
			object = managedSystem[i].getObjects(VSCSI);
			if (object != null) {
				for (j=0; j<object.length; j++) {
					
					col = 0;
					
					s = object[j].getVarValues("lpar_name");
					n = ds.addLabel(col, row, s,0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
					if (n>size[col]) size[col]=n; col++;
					
					s = object[j].getVarValues("slot_num");
					ds.addInteger(col, row, s,0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
					n=6; if (n>size[col]) size[col]=n; col++;

					
					s = object[j].getVarValues("state");
					if (s!=null) {
						if (s[0].equals("1"))
							n = ds.addLabel(col, row, "On", B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
						else
							n = ds.addLabel(col, row, "Off", B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
						if (n>size[col]) size[col]=n; col++;
					} else
						col++;
					
					
					s = object[j].getVarValues("is_required");
					if (s!=null) {
						if (s[0].equals("1"))
							n = ds.addLabel(col, row, "True", B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
						else
							n = ds.addLabel(col, row, "False", B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
						if (n>size[col]) size[col]=n; col++;
					} else
						col++;
					
					s = object[j].getVarValues("adapter_type");
					n = ds.addLabel(col, row, s,0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
					if (n>size[col]) size[col]=n; col++;
					
					s = object[j].getVarValues("remote_lpar_name");
					n = ds.addLabel(col, row, s,0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
					if (n>size[col]) size[col]=n; col++;
					
					s = object[j].getVarValues("remote_slot_num");
					if (s!=null && s[0].equals("any"))
						n = ds.addLabel(col, row, s,0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
					else
						ds.addInteger(col, row, s,0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
					n=6; if (n>size[col]) size[col]=n; col++;
					
					n = ds.addLabel(col, row, managedSystem[i].getVarValues("name"),0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
					if (n>size[col]) size[col]=n; col++;
					n = ds.addLabel(col, row, managedSystem[i].getVarValues("serial_num"),0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
					if (n>size[col]) size[col]=n; col++;
												
					row++;				
				}				
			}
		}
		
		for (i=0; i<size.length; i++)
			ds.setColSize(i, size[i]+2);

		return ds;
	}
	
	
	
	private void createHMCExcel(WritableSheet sheet) {
		DataSheet ds = createHMCSheet();
		if (ds!=null)
			ds.createExcelSheet(sheet);
	}
	
	private void createHMCHTML(String fileName) {
		DataSheet ds = createHMCSheet();
		if (ds!=null) {
			ds.createHTMLSheet(fileName);
			addButton("HMC",new File(fileName).getName());
		}
	}
	
	private void createHMCCSV(String fileName) {
		DataSheet ds = createHMCSheet();
		if (ds!=null) {
			ds.setSeparator(csvSeparator);
			ds.createCSVSheet(fileName);
		}
	}
	
	private DataSheet createHMCSheet() {
		int row;
		int i;
		String s[];
		
		DataSheet ds = new DataSheet();
					
		row = 0;
		
		ds.mergeCells(0, row, 1, row);
		ds.addLabel(0,row,"Hardware",BOLD|CENTRE|VCENTRE|B_ALL_MED);
		row++;
		
		ds.addLabel(0,row,"BIOS",BOLD|B_ALL_MED);
		ds.addLabel(1,row,hmc.getVarValues("bios"),0,B_ALL_LOW);
		row++;
		
		ds.addLabel(0,row,"Model",BOLD|B_ALL_MED);
		ds.addLabel(1,row,hmc.getVarValues("model"),0,B_ALL_LOW);
		row++;
		
		ds.addLabel(0,row,"Serial",BOLD|B_ALL_MED);
		ds.addLabel(1,row,hmc.getVarValues("serial"),0,B_ALL_LOW);
		row++;
		
		
		
		row=0;
		
		ds.mergeCells(3, row, 4, row);
		ds.addLabel(3,row,"Software",BOLD|CENTRE|VCENTRE|B_ALL_MED);
		row++;
		
		ds.addLabel(3,row,"Version",BOLD|B_ALL_MED);
		ds.addLabel(4,row,hmc.getVarValues("version"),0,B_ALL_LOW);
		row++;
		
		ds.addLabel(3,row,"Release",BOLD|B_ALL_MED);
		ds.addLabel(4,row,hmc.getVarValues("release"),0,B_ALL_LOW);
		row++;
		
		ds.addLabel(3,row,"Service Pack",BOLD|B_ALL_MED);
		ds.addLabel(4,row,hmc.getVarValues("sp"),0,B_ALL_LOW);
		row++;
		
		ds.addLabel(3,row,"Build Level",BOLD|B_ALL_MED);
		ds.addLabel(4,row,hmc.getVarValues("build_level"),0,B_ALL_LOW);
		row++;
		
		ds.addLabel(3,row,"Base Version",BOLD|B_ALL_MED);
		ds.addLabel(4,row,hmc.getVarValues("base_version"),0,B_ALL_LOW);
		row++;
		
		ds.addLabel(3,row,"Fixes",BOLD|B_ALL_MED);
		s = hmc.getVarValues("fixes");
		if (s!=null)
			for (i=0; i<s.length; i++) {
				ds.addLabel(4+i,row,s[i],B_ALL_LOW);
			}
		row++;
		
		
		row=15;
		
		
		ds.mergeCells(0, row, 6, row);
		ds.addLabel(0,row,"Network",BOLD|CENTRE|VCENTRE|B_ALL_MED);
		row++;
		
		ds.addLabel(0,row,"Hostname",BOLD|B_ALL_MED);
		ds.addLabel(1,row,hmc.getVarValues("hostname"),0,B_ALL_LOW);
		row++;
		
		ds.addLabel(0,row,"Domain",BOLD|B_ALL_MED);
		ds.addLabel(1,row,hmc.getVarValues("domain"),0,B_ALL_LOW);
		row++;
		
		ds.addLabel(0,row,"Gateway",BOLD|B_ALL_MED);
		ds.addLabel(1,row,hmc.getVarValues("gateway"),0,B_ALL_LOW);
		row++;
		
		ds.addLabel(0,row,"Name Server",BOLD|B_ALL_MED);
		ds.addLabel(1,row,hmc.getVarValues("nameserver"),0,B_ALL_LOW);
		row++;
		
		ds.addLabel(0,row,"DNS suffix",BOLD|B_ALL_MED);
		ds.addLabel(1,row,hmc.getVarValues("domainsuffix"),0,B_ALL_LOW);
		row++;
		
		row++;
		for (i=1; i<=6; i++)
			ds.addLabel(i,row,"eth"+(i-1),BOLD|B_ALL_MED);
		row++;
		
		ds.addLabel(0,row,	"IPv4 addr",BOLD|B_ALL_MED);
		ds.addLabel(0,row+1,	"IPv4 netmask",BOLD|B_ALL_MED);
		ds.addLabel(0,row+2,	"IPv4 dhcp",BOLD|B_ALL_MED);
		ds.addLabel(0,row+3,	"IPv6 addr",BOLD|B_ALL_MED);
		ds.addLabel(0,row+4,	"IPv6 auto",BOLD|B_ALL_MED);
		ds.addLabel(0,row+5,	"IPv6 privacy",BOLD|B_ALL_MED);
		ds.addLabel(0,row+6,	"IPv6 dhcp",BOLD|B_ALL_MED);
		ds.addLabel(0,row+7,	"LPAR comm",BOLD|B_ALL_MED);
		ds.addLabel(0,row+8,	"Jumbo frame",BOLD|B_ALL_MED);
		ds.addLabel(0,row+9,	"Speed",BOLD|B_ALL_MED);
		ds.addLabel(0,row+10,"Duplex",BOLD|B_ALL_MED);
		ds.addLabel(0,row+11,"TSO",BOLD|B_ALL_MED);
		
		for (i=1; i<=6; i++) {
			ds.addLabel(i,row,	hmc.getVarValues("ipv4addr_eth"+(i-1)),0,B_ALL_LOW);
			ds.addLabel(i,row+1,	hmc.getVarValues("ipv4netmask_eth"+(i-1)),0,B_ALL_LOW);
			ds.addLabel(i,row+2,	hmc.getVarValues("ipv4dhcp_eth"+(i-1)),0,B_ALL_LOW);
			ds.addLabel(i,row+3,	hmc.getVarValues("ipv6addr_eth"+(i-1)),0,B_ALL_LOW);
			ds.addLabel(i,row+4,	hmc.getVarValues("ipv6auto_eth"+(i-1)),0,B_ALL_LOW);
			ds.addLabel(i,row+5,	hmc.getVarValues("ipv6privacy_eth"+(i-1)),0,B_ALL_LOW);
			ds.addLabel(i,row+6,	hmc.getVarValues("ipv6dhcp_eth"+(i-1)),0,B_ALL_LOW);
			ds.addLabel(i,row+7,	hmc.getVarValues("lparcomm_eth"+(i-1)),0,B_ALL_LOW);
			ds.addLabel(i,row+8,	hmc.getVarValues("jumboframe_eth"+(i-1)),0,B_ALL_LOW);
			ds.addLabel(i,row+9,	hmc.getVarValues("speed_eth"+(i-1)),0,B_ALL_LOW);
			ds.addLabel(i,row+10,hmc.getVarValues("duplex_eth"+(i-1)),0,B_ALL_LOW);
			ds.addLabel(i,row+11,hmc.getVarValues("tso_eth"+(i-1)),0,B_ALL_LOW);			
		}
		


		for (i=0; i<10; i++)
			ds.setColSize(i, 28);	
		
		return ds;
	}
	
	
	
	private void createHMCSheet(WritableSheet sheet)  throws RowsExceededException, WriteException {
		int row;
		int i,j;
		String s[];
		
			
		row = 0;
		
		sheet.mergeCells(0, row, 1, row);
		addLabel(sheet,0,row,"Hardware",formatLabel(BOLD|CENTRE|VCENTRE|B_ALL_MED));
		row++;
		
		addLabel(sheet,0,row,"BIOS",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,1,row,hmc.getVarValues("bios"),0,formatLabel(B_ALL_LOW));
		row++;
		
		addLabel(sheet,0,row,"Model",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,1,row,hmc.getVarValues("model"),0,formatLabel(B_ALL_LOW));
		row++;
		
		addLabel(sheet,0,row,"Serial",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,1,row,hmc.getVarValues("serial"),0,formatLabel(B_ALL_LOW));
		row++;
		
		
		
		row=0;
		
		sheet.mergeCells(3, row, 4, row);
		addLabel(sheet,3,row,"Software",formatLabel(BOLD|CENTRE|VCENTRE|B_ALL_MED));
		row++;
		
		addLabel(sheet,3,row,"Version",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,4,row,hmc.getVarValues("version"),0,formatLabel(B_ALL_LOW));
		row++;
		
		addLabel(sheet,3,row,"Release",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,4,row,hmc.getVarValues("release"),0,formatLabel(B_ALL_LOW));
		row++;
		
		addLabel(sheet,3,row,"Service Pack",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,4,row,hmc.getVarValues("sp"),0,formatLabel(B_ALL_LOW));
		row++;
		
		addLabel(sheet,3,row,"Build Level",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,4,row,hmc.getVarValues("build_level"),0,formatLabel(B_ALL_LOW));
		row++;
		
		addLabel(sheet,3,row,"Base Version",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,4,row,hmc.getVarValues("base_version"),0,formatLabel(B_ALL_LOW));
		row++;
		
		addLabel(sheet,3,row,"Fixes",formatLabel(BOLD|B_ALL_MED));
		s = hmc.getVarValues("fixes");
		if (s!=null)
			for (i=0; i<s.length; i++) {
				addLabel(sheet,4+i,row,s[i],formatLabel(B_ALL_LOW));
			}
		row++;
		
		
		row=15;
		
		
		sheet.mergeCells(0, row, 6, row);
		addLabel(sheet,0,row,"Network",formatLabel(BOLD|CENTRE|VCENTRE|B_ALL_MED));
		row++;
		
		addLabel(sheet,0,row,"Hostname",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,1,row,hmc.getVarValues("hostname"),0,formatLabel(B_ALL_LOW));
		row++;
		
		addLabel(sheet,0,row,"Domain",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,1,row,hmc.getVarValues("domain"),0,formatLabel(B_ALL_LOW));
		row++;
		
		addLabel(sheet,0,row,"Gateway",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,1,row,hmc.getVarValues("gateway"),0,formatLabel(B_ALL_LOW));
		row++;
		
		addLabel(sheet,0,row,"Name Server",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,1,row,hmc.getVarValues("nameserver"),0,formatLabel(B_ALL_LOW));
		row++;
		
		addLabel(sheet,0,row,"DNS suffix",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,1,row,hmc.getVarValues("domainsuffix"),0,formatLabel(B_ALL_LOW));
		row++;
		
		row++;
		for (i=1; i<=6; i++)
			addLabel(sheet,i,row,"eth"+(i-1),formatLabel(BOLD|B_ALL_MED));
		row++;
		
		addLabel(sheet,0,row,	"IPv4 addr",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,0,row+1,	"IPv4 netmask",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,0,row+2,	"IPv4 dhcp",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,0,row+3,	"IPv6 addr",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,0,row+4,	"IPv6 auto",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,0,row+5,	"IPv6 privacy",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,0,row+6,	"IPv6 dhcp",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,0,row+7,	"LPAR comm",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,0,row+8,	"Jumbo frame",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,0,row+9,	"Speed",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,0,row+10,"Duplex",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,0,row+11,"TSO",formatLabel(BOLD|B_ALL_MED));
		
		for (i=1; i<=6; i++) {
			addLabel(sheet,i,row,	hmc.getVarValues("ipv4addr_eth"+(i-1)),0,formatLabel(B_ALL_LOW));
			addLabel(sheet,i,row+1,	hmc.getVarValues("ipv4netmask_eth"+(i-1)),0,formatLabel(B_ALL_LOW));
			addLabel(sheet,i,row+2,	hmc.getVarValues("ipv4dhcp_eth"+(i-1)),0,formatLabel(B_ALL_LOW));
			addLabel(sheet,i,row+3,	hmc.getVarValues("ipv6addr_eth"+(i-1)),0,formatLabel(B_ALL_LOW));
			addLabel(sheet,i,row+4,	hmc.getVarValues("ipv6auto_eth"+(i-1)),0,formatLabel(B_ALL_LOW));
			addLabel(sheet,i,row+5,	hmc.getVarValues("ipv6privacy_eth"+(i-1)),0,formatLabel(B_ALL_LOW));
			addLabel(sheet,i,row+6,	hmc.getVarValues("ipv6dhcp_eth"+(i-1)),0,formatLabel(B_ALL_LOW));
			addLabel(sheet,i,row+7,	hmc.getVarValues("lparcomm_eth"+(i-1)),0,formatLabel(B_ALL_LOW));
			addLabel(sheet,i,row+8,	hmc.getVarValues("jumboframe_eth"+(i-1)),0,formatLabel(B_ALL_LOW));
			addLabel(sheet,i,row+9,	hmc.getVarValues("speed_eth"+(i-1)),0,formatLabel(B_ALL_LOW));
			addLabel(sheet,i,row+10,hmc.getVarValues("duplex_eth"+(i-1)),0,formatLabel(B_ALL_LOW));
			addLabel(sheet,i,row+11,hmc.getVarValues("tso_eth"+(i-1)),0,formatLabel(B_ALL_LOW));			
		}
		


		for (i=0; i<10; i++)
			sheet.setColumnView(i, 28);			
	}
	
	private String lparid_to_name(GenericData ms, String id) {
		if (id.equals("0x00000000"))
			return("");
		
		GenericData object[] = ms.getObjects(CONFIG_LPAR);
		int lpar_id = Integer.parseInt(id.substring(2), 16);
		int i=0;
		int n;
		
		while (i<object.length) {
			n = Integer.parseInt(object[i].getVarValues("lpar_id")[0]);
			if (n==lpar_id)
				return object[i].getVarValues("name")[0];
			i++;
		}
		
		return("ERROR");		
	}
	
	
	private void createVSCSIMapSheet(WritableSheet sheet)  throws RowsExceededException, WriteException {
		GenericData object[];
		GenericData hdisk[];
		int row;
		int i,j,k,v;
		String s[];		
		int lines=0;
		int size[]=new int[11];
		int n;
		String device;
		String vios;
		String s2[];
		
			
		row = 0;
		for (i=0; i<size.length; i++)
			size[i] = 0;
		
		for (i=0; i<managedSystem.length; i++) {
			
			/*
			 * Show start of system
			 */
			s = managedSystem[i].getVarValues("name");
			sheet.mergeCells(0, row, 10, row+1);
			addLabel(sheet,0,row,s[0],formatLabel(BOLD|CENTRE|VCENTRE|GREEN));
			row++;
			row++;
			
			/*
			 * Setup titles for virtual slots
			 */ 			
			n=addLabel(sheet,0,row,"VIOS",formatLabel(BOLD|CENTRE|B_ALL_MED));
			if (n>size[0]) size[0]=n;
			n=addLabel(sheet,1,row,"Slot",formatLabel(BOLD|CENTRE|B_ALL_MED));
			if (n>size[1]) size[1]=n;
			n=addLabel(sheet,2,row,"Name",formatLabel(BOLD|CENTRE|B_ALL_MED));
			if (n>size[2]) size[2]=n;
			n=addLabel(sheet,3,row,"Client",formatLabel(BOLD|CENTRE|B_ALL_MED));
			if (n>size[3]) size[3]=n;
			n=addLabel(sheet,4,row,"VTD",formatLabel(BOLD|CENTRE|B_ALL_MED));
			if (n>size[4]) size[4]=n;
			n=addLabel(sheet,5,row,"Status",formatLabel(BOLD|CENTRE|B_ALL_MED));
			if (n>size[5]) size[5]=n;
			n=addLabel(sheet,6,row,"LUN",formatLabel(BOLD|CENTRE|B_ALL_MED));
			if (n>size[6]) size[6]=n;
			n=addLabel(sheet,7,row,"Backing device",formatLabel(BOLD|CENTRE|B_ALL_MED));
			if (n>size[7]) size[7]=n;
			n=addLabel(sheet,8,row,"Phys Loc",formatLabel(BOLD|CENTRE|B_ALL_MED));
			if (n>size[8]) size[8]=n;
			n=addLabel(sheet,9,row,"Mirrored",formatLabel(BOLD|CENTRE|B_ALL_MED));
			if (n>size[9]) size[9]=n;
			n=addLabel(sheet,10,row,"Identifier",formatLabel(BOLD|CENTRE|B_ALL_MED));
			if (n>size[10]) size[10]=n;
			
			row ++;
			
			/*
			 * Write virtual slot variables values
			 */
			hdisk = managedSystem[i].getObjects(HDISK);
			object = managedSystem[i].getObjects(VSCSIMAP);
			if (object != null) {
				for (j=0; j<object.length; j++) {
					
					lines = object[j].getVarValues("VTD").length;
					
					s = object[j].getVarValues("VIOS");
					sheet.mergeCells(0, row, 0, row+lines-1);
					n=addLabel(sheet, 0, row, s,0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					if (n>size[0]) size[0]=n; 
					vios=s[0];
					
					s = object[j].getVarValues("slot");
					sheet.mergeCells(1, row, 1, row+lines-1);
					addNumber(sheet, 1, row, s,0, formatInt(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					
					s = object[j].getVarValues("SVSA");
					sheet.mergeCells(2, row, 2, row+lines-1);
					n=addLabel(sheet, 2, row, s,0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					if (n>size[2]) size[2]=n; 
					
					s = object[j].getVarValues("client");
					sheet.mergeCells(3, row, 3, row+lines-1);
					//n=addLabel(sheet, 3, row, s,0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					n=addLabel(sheet, 3, row, s[0]+" \""+lparid_to_name(managedSystem[i],s[0])+"\"", formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					if (n>size[3]) size[3]=n; 
					
					s = object[j].getVarValues("VTD");
					for (k=0; k<lines; k++) {
						n=addLabel(sheet, 4, row+k, object[j].getVarValues("VTD")				,k, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
						if (n>size[4]) size[4]=n; 
						n=addLabel(sheet, 5, row+k, object[j].getVarValues("Status")			,k, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
						if (n>size[5]) size[5]=n; 
						n=addLabel(sheet, 6, row+k, object[j].getVarValues("LUN")				,k, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
						if (n>size[6]) size[6]=n;
						
						n=addLabel(sheet, 7, row+k, object[j].getVarValues("BackingDevice")	,k, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
						if (n>size[7]) size[7]=n; 
						device = object[j].getVarValues("BackingDevice")[k];
						
						n=addLabel(sheet, 8, row+k, object[j].getVarValues("physloc")			,k, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
						if (n>size[8]) size[8]=n; 
						n=addLabel(sheet, 9, row+k, object[j].getVarValues("mirror")			,k, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
						if (n>size[9]) size[9]=n; 
						
						if (device!=null && device.startsWith("hdisk") && hdisk!=null) {
							for (v=0; v<hdisk.length; v++) {
								s2 = hdisk[v].getVarValues("VIOS");
								if (s2!=null && s2[0]!=null && s2[0].equals(vios)) {
									s2 = hdisk[v].getVarValues(device);
									if (s2!=null) {
										n = addLabel(sheet, 10, row+k, s2, 0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
										if (n>size[10]) size[10]=n; 
									}
								}
							}
						}
					}
								
					row+=s.length;				
				}				
			}
			
			
			row +=2;
		}

		for (i=0; i<size.length; i++)
			sheet.setColumnView(i, size[i]+2);
	}
	

	
	private void createVSCSIMapSheetRowBasedExcel(WritableSheet sheet) {
		DataSheet ds = createVSCSIMapSheetRowBased();
		if (ds!=null)
			ds.createExcelSheet(sheet);
	}
	
	private void createVSCSIMapSheetRowBasedHTML(String fileName) {
		DataSheet ds = createVSCSIMapSheetRowBased();
		if (ds!=null) {
			ds.createHTMLSheet(fileName);
			addButton("VSCSI Map",new File(fileName).getName());
		}
	}
	
	private void createVSCSIMapSheetRowBasedCSV(String fileName) {
		DataSheet ds = createVSCSIMapSheetRowBased();
		if (ds!=null) {
			ds.setSeparator(csvSeparator);
			ds.createCSVSheet(fileName);
		}
	}
	
	
	private DataSheet createVSCSIMapSheetRowBased() {
		DataSheet ds = new DataSheet();
		GenericData object[];
		GenericData hdisk[];
		int row, col;
		int i,j,k,v;		
		int lines=0;
		int size[]=new int[13];
		int n;
		String device;
		String vios;
		String s2[];
		
		String sysName[], serial[];
		String slot[], svsa[], client[];
		
			
		row = 0;
		col = 0;
		for (i=0; i<size.length; i++)
			size[i] = 0;
		
		
		/*
		 * Setup titles for virtual slots
		 */ 			
		n=ds.addLabel(col,row,"VIOS",BOLD|CENTRE|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n=ds.addLabel(col,row,"Slot",BOLD|CENTRE|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n=ds.addLabel(col,row,"Name",BOLD|CENTRE|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n=ds.addLabel(col,row,"Client",BOLD|CENTRE|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n=ds.addLabel(col,row,"VTD",BOLD|CENTRE|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n=ds.addLabel(col,row,"Status",BOLD|CENTRE|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n=ds.addLabel(col,row,"LUN",BOLD|CENTRE|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n=ds.addLabel(col,row,"Backing device",BOLD|CENTRE|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n=ds.addLabel(col,row,"Phys Loc",BOLD|CENTRE|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n=ds.addLabel(col,row,"Mirrored",BOLD|CENTRE|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n=ds.addLabel(col,row,"Identifier",BOLD|CENTRE|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n=ds.addLabel(col,row,"Managed System Name",BOLD|CENTRE|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n=ds.addLabel(col,row,"Managed System Serial",BOLD|CENTRE|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		
		row ++;
		
		for (i=0; i<managedSystem.length; i++) {
			
			
			/*
			 * Write virtual slot variables values
			 */
			hdisk = managedSystem[i].getObjects(HDISK);
			object = managedSystem[i].getObjects(VSCSIMAP);
			if (object != null) {
				for (j=0; j<object.length; j++) {
					
					col = 0;
					
					lines = object[j].getVarValues("VTD").length;
					
					vios = object[j].getVarValues("VIOS")[0];
					for (k=0; k<lines; k++) {
						n=ds.addLabel(col, row+k, vios, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
						if (n>size[col]) size[col]=n; 
					}
					col++;
				
					slot = object[j].getVarValues("slot");
					for (k=0; k<lines; k++) {
						ds.addInteger(col, row+k, slot,0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
						n=5; if (n>size[col]) size[col]=n; 
					}
					col++;
					
					svsa = object[j].getVarValues("SVSA");
					for (k=0; k<lines; k++) {
						n=ds.addLabel(col, row+k, svsa,0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
						if (n>size[col]) size[col]=n; 
					}
					col++;
					
					client = object[j].getVarValues("client");
					for (k=0; k<lines; k++) {
						n=ds.addLabel(col, row+k, client[0]+" \""+lparid_to_name(managedSystem[i],client[0])+"\"", VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
						if (n>size[col]) size[col]=n; 
					}
					col++;
					
					for (k=0; k<lines; k++) {
						n=ds.addLabel(col, row+k, object[j].getVarValues("VTD")		,k, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; 
						n=ds.addLabel(col+1, row+k, object[j].getVarValues("Status")	,k, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col+1]) size[col+1]=n; 
						n=ds.addLabel(col+2, row+k, object[j].getVarValues("LUN")		,k, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col+2]) size[col+2]=n; 						
						n=ds.addLabel(col+3, row+k, object[j].getVarValues("BackingDevice")	,k, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col+3]) size[col+3]=n; 
						
						device = object[j].getVarValues("BackingDevice")[k];
						
						n=ds.addLabel(col+4, row+k, object[j].getVarValues("physloc")	,k, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col+4]) size[col+4]=n; 
						n=ds.addLabel(col+5, row+k, object[j].getVarValues("mirror")	,k, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col+5]) size[col+5]=n; 
						
						if (device!=null && device.startsWith("hdisk") && hdisk!=null) {
							for (v=0; v<hdisk.length; v++) {
								s2 = hdisk[v].getVarValues("VIOS");
								if (s2!=null && s2[0]!=null && s2[0].equals(vios)) {
									s2 = hdisk[v].getVarValues(device);
									if (s2!=null) {
										n = ds.addLabel(col+6, row+k, s2, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col+6]) size[col+6]=n; 
									}
								}
							}
						}
					}
					col = col + 7;
					
					sysName = managedSystem[i].getVarValues("name");
					serial = managedSystem[i].getVarValues("serial_num");
					for (k=0; k<lines; k++) {
						n = ds.addLabel(col, row+k, sysName,0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n;
					}
					col++;
					for (k=0; k<lines; k++) {
						n = ds.addLabel(col, row+k, serial,0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n;
					}
					col++;
								
					row+=lines;				
				}				
			}
		}

		for (i=0; i<size.length; i++)
			ds.setColSize(i, size[i]+2);

		return ds;
	}
	
	
	
	
	private void createVSCSIMapSheetRowBased(WritableSheet sheet)  throws RowsExceededException, WriteException {
		GenericData object[];
		GenericData hdisk[];
		int row;
		int i,j,k,v;		
		int lines=0;
		int size[]=new int[13];
		int n;
		String device;
		String vios;
		String s2[];
		
		String sysName[], serial[];
		String slot[], svsa[], client[];
		
			
		row = 0;
		for (i=0; i<size.length; i++)
			size[i] = 0;
		
		
		/*
		 * Setup titles for virtual slots
		 */ 			
		n=addLabel(sheet,0,row,"VIOS",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN));
		if (n>size[0]) size[0]=n;
		n=addLabel(sheet,1,row,"Slot",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN));
		if (n>size[1]) size[1]=n;
		n=addLabel(sheet,2,row,"Name",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN));
		if (n>size[2]) size[2]=n;
		n=addLabel(sheet,3,row,"Client",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN));
		if (n>size[3]) size[3]=n;
		n=addLabel(sheet,4,row,"VTD",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN));
		if (n>size[4]) size[4]=n;
		n=addLabel(sheet,5,row,"Status",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN));
		if (n>size[5]) size[5]=n;
		n=addLabel(sheet,6,row,"LUN",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN));
		if (n>size[6]) size[6]=n;
		n=addLabel(sheet,7,row,"Backing device",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN));
		if (n>size[7]) size[7]=n;
		n=addLabel(sheet,8,row,"Phys Loc",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN));
		if (n>size[8]) size[8]=n;
		n=addLabel(sheet,9,row,"Mirrored",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN));
		if (n>size[9]) size[9]=n;
		n=addLabel(sheet,10,row,"Identifier",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN));
		if (n>size[10]) size[10]=n;
		n=addLabel(sheet,11,row,"Managed System Name",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN));
		if (n>size[11]) size[11]=n;
		n=addLabel(sheet,12,row,"Managed System Serial",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN));
		if (n>size[12]) size[12]=n;
		
		row ++;
		
		for (i=0; i<managedSystem.length; i++) {
			
			
			/*
			 * Write virtual slot variables values
			 */
			hdisk = managedSystem[i].getObjects(HDISK);
			object = managedSystem[i].getObjects(VSCSIMAP);
			if (object != null) {
				for (j=0; j<object.length; j++) {
					
					lines = object[j].getVarValues("VTD").length;
					
					vios = object[j].getVarValues("VIOS")[0];
					for (k=0; k<lines; k++) 
						n=addLabel(sheet, 0, row+k, vios, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					if (n>size[0]) size[0]=n; 
				
					slot = object[j].getVarValues("slot");
					for (k=0; k<lines; k++) 
						addNumber(sheet, 1, row+k, slot,0, formatInt(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					
					svsa = object[j].getVarValues("SVSA");
					for (k=0; k<lines; k++) 
						n=addLabel(sheet, 2, row+k, svsa,0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					if (n>size[2]) size[2]=n; 
					
					client = object[j].getVarValues("client");
					for (k=0; k<lines; k++) 
						n=addLabel(sheet, 3, row+k, client[0]+" \""+lparid_to_name(managedSystem[i],client[0])+"\"", formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					if (n>size[3]) size[3]=n; 
					
					for (k=0; k<lines; k++) {
						n=addLabel(sheet, 4, row+k, object[j].getVarValues("VTD")				,k, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
						if (n>size[4]) size[4]=n; 
						n=addLabel(sheet, 5, row+k, object[j].getVarValues("Status")			,k, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
						if (n>size[5]) size[5]=n; 
						n=addLabel(sheet, 6, row+k, object[j].getVarValues("LUN")				,k, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
						if (n>size[6]) size[6]=n;
						
						n=addLabel(sheet, 7, row+k, object[j].getVarValues("BackingDevice")	,k, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
						if (n>size[7]) size[7]=n; 
						
						device = object[j].getVarValues("BackingDevice")[k];
						
						n=addLabel(sheet, 8, row+k, object[j].getVarValues("physloc")			,k, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
						if (n>size[8]) size[8]=n; 
						n=addLabel(sheet, 9, row+k, object[j].getVarValues("mirror")			,k, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
						if (n>size[9]) size[9]=n; 
						
						if (device!=null && device.startsWith("hdisk") && hdisk!=null) {
							for (v=0; v<hdisk.length; v++) {
								s2 = hdisk[v].getVarValues("VIOS");
								if (s2!=null && s2[0]!=null && s2[0].equals(vios)) {
									s2 = hdisk[v].getVarValues(device);
									if (s2!=null) {
										n = addLabel(sheet, 10, row+k, s2, 0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
										if (n>size[10]) size[10]=n; 
									}
								}
							}
						}
					}
					
					sysName = managedSystem[i].getVarValues("name");
					serial = managedSystem[i].getVarValues("serial_num");
					for (k=0; k<lines; k++) 
						n = addLabel(sheet, 11, row+k, sysName,0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					if (n>size[11]) size[11]=n;
					for (k=0; k<lines; k++) 
						n = addLabel(sheet, 12, row+k, serial,0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					if (n>size[12]) size[12]=n;
								
					row+=lines;				
				}				
			}
		}

		for (i=0; i<size.length; i++)
			sheet.setColumnView(i, size[i]+2);
	}
	
	
	private void createSEASheetExcel(WritableSheet sheet) {
		DataSheet ds = createSEASheet();
		if (ds!=null)
			ds.createExcelSheet(sheet);
	}
	
	private void createSEASheetHTML(String fileName) {
		DataSheet ds = createSEASheet();
		if (ds!=null) {
			ds.createHTMLSheet(fileName);
			addButton("SEA",new File(fileName).getName());
		}
	}
	
	private void createSEASheetCSV(String fileName) {
		DataSheet ds = createSEASheet();
		if (ds!=null) {
			ds.setSeparator(csvSeparator);
			ds.createCSVSheet(fileName);
		}
	}

	private DataSheet createSEASheet() {
		DataSheet ds = new DataSheet();
		GenericData object[];
		int row,col;
		int i,j,k;		
		int size[] = new int[1000];
		int n;
		String names[];
		int maxVirt = 0;
		GenericData eth[];
		GenericData myEth;
		String vios;
		
		
		for (i=0; i<size.length; i++)
			size[i] = 4;
		
			
		row = 0;
		
		
		// Compute max virtual adapters in all SEAs
		for (i=0; i<managedSystem.length; i++) {
			object = managedSystem[i].getObjects(SEA);
			for (j=0; object!=null && j<object.length; j++) {
				names = object[j].getVarValues("virt_adapters");
				if (names.length>maxVirt)
					maxVirt=names.length;
			}
		}
		
		
		
		/*
		 * Setup titles 
		 */ 	
		col=0;
		n=ds.addLabel(col,row,"Managed System Name",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45); col++;
		n=ds.addLabel(col,row,"Managed System Serial",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"VIOS",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"SEA",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45); col++;
		n=ds.addLabel(col,row,"ctl_chan",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45); col++;
		n=ds.addLabel(col,row,"ctl_chan slot",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45); col++;
		n=ds.addLabel(col,row,"gvrp",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45); col++;
		n=ds.addLabel(col,row,"ha_mode",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45); col++;
		n=ds.addLabel(col,row,"jumbo_frames",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45); col++;
		n=ds.addLabel(col,row,"large_receive",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"largesend",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"netaddr",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"accounting",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"lldpsvc",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"qos_mode",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"thread",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"pvid",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"pvid_adapter",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"real_adapter",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"real_adapter slot",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"thread",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		
		for (k=1;k<=maxVirt; k++) {
			n=ds.addLabel(col,row,"Virt#"+k,BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45); col++;
			n=ds.addLabel(col,row,"Virt#"+k+" slot",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		}
			
		
		row ++;
		
		for (i=0; i<managedSystem.length; i++) {
			
			
			/*
			 * Write virtual slot variables values
			 */
			eth = managedSystem[i].getObjects(ETH);
			object = managedSystem[i].getObjects(SEA);
			if (object != null) {
				for (j=0; j<object.length; j++) {
										
					// Get eth mapping object
					vios = object[j].getVarValues("VIOS")[0];
					for (k=0, myEth=null; k<eth.length; k++) {
						myEth = eth[k];
						if (myEth.getVarValues("VIOS")[0].equals(vios))
							break;
					}
					
					col=0;
					n=ds.addLabel( col, row, managedSystem[i].getVarValues("name"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, managedSystem[i].getVarValues("serial_num"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, object[j].getVarValues("VIOS"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;					
					n=ds.addLabel( col, row, object[j].getVarValues("SEA"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;				
					n=ds.addLabel( col, row, object[j].getVarValues("ctl_chan"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					
					n=ds.addLabel( col, row, myEth.getVarValues(object[j].getVarValues("ctl_chan")[0]), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					
					n=ds.addLabel( col, row, object[j].getVarValues("gvrp"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, object[j].getVarValues("ha_mode"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, object[j].getVarValues("jumbo_frames"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, object[j].getVarValues("large_receive"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, object[j].getVarValues("largesend"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, object[j].getVarValues("netaddr"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, object[j].getVarValues("accounting"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, object[j].getVarValues("lldpsvc"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, object[j].getVarValues("qos_mode"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, object[j].getVarValues("thread"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, object[j].getVarValues("pvid"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, object[j].getVarValues("pvid_adapter"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, object[j].getVarValues("real_adapter"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					
					n=ds.addLabel( col, row, myEth.getVarValues(object[j].getVarValues("real_adapter")[0]), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					
					n=ds.addLabel( col, row, object[j].getVarValues("thread"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;

					names = object[j].getVarValues("virt_adapters");
					for (k=0; k<maxVirt; k++) {
						if (k<names.length) {
							n=ds.addLabel( col, row, names, k, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
							n=ds.addLabel( col, row, myEth.getVarValues(names[k]), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
						} else
							col+=2;
					}
								
					row++;				
				}				
			}
		}
		
		row++;
		row++;
		row++;
		
		
		// Compute max virtual adapters in all ETherchannel
		for (i=0; i<managedSystem.length; i++) {
			object = managedSystem[i].getObjects(ETHERCHANNEL);
			for (j=0; object!=null && j<object.length; j++) {
				names = object[j].getVarValues("adapter_names");
				if (names.length>maxVirt)
					maxVirt=names.length;
			}
		}
		
		/*
		 * Setup titles for ETherchannel
		 */ 	
		col=0;
		n=ds.addLabel(col,row,"Managed System Name",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"Managed System Serial",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"VIOS",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"Etherchannel",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45); col++;
		n=ds.addLabel(col,row,"alt_addr",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45); col++;
		n=ds.addLabel(col,row,"auto_recovery",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45); col++;
		n=ds.addLabel(col,row,"hash_mode",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45); col++;
		n=ds.addLabel(col,row,"interval",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45); col++;
		n=ds.addLabel(col,row,"mode",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45); col++;
		n=ds.addLabel(col,row,"netaddr",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45); col++;
		n=ds.addLabel(col,row,"noloss_failover",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"num_retries",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"retry_time",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"use_alt_addr",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"use_jumbo_frame",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"backup_adapter",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"backup_adapter slot",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		
		
		for (k=1;k<=maxVirt; k++) {
			n=ds.addLabel(col,row,"Adapter#"+k,BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
			n=ds.addLabel(col,row,"Adapter#"+k+" slot",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		}
			
		
		row ++;
		
		for (i=0; i<managedSystem.length; i++) {
			
			
			/*
			 * Write virtual slot variables values
			 */
			eth = managedSystem[i].getObjects(ETH);
			object = managedSystem[i].getObjects(ETHERCHANNEL);
			if (object != null) {
				for (j=0; j<object.length; j++) {
										
					// Get eth mapping object
					vios = object[j].getVarValues("VIOS")[0];
					for (k=0, myEth=null; k<eth.length; k++) {
						myEth = eth[k];
						if (myEth.getVarValues("VIOS")[0].equals(vios))
							break;
					}
					
					col=0;
					n=ds.addLabel( col, row, managedSystem[i].getVarValues("name"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, managedSystem[i].getVarValues("serial_num"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, object[j].getVarValues("VIOS"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;					
					n=ds.addLabel( col, row, object[j].getVarValues("ETHERCHANNEL"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;				
					n=ds.addLabel( col, row, object[j].getVarValues("alt_addr"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;					
					n=ds.addLabel( col, row, object[j].getVarValues("auto_recovery"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, object[j].getVarValues("hash_mode"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;					
					n=ds.addLabel( col, row, object[j].getVarValues("interval"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;					
					n=ds.addLabel( col, row, object[j].getVarValues("mode"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, object[j].getVarValues("netaddr"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, object[j].getVarValues("noloss_failover"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, object[j].getVarValues("num_retries"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, object[j].getVarValues("retry_time"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, object[j].getVarValues("use_alt_addr"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, object[j].getVarValues("use_jumbo_frame"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, object[j].getVarValues("backup_adapter"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					
					n=ds.addLabel( col, row, myEth.getVarValues(object[j].getVarValues("backup_adapter")[0]), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					
					names = object[j].getVarValues("adapter_names");
					for (k=0; k<maxVirt; k++) {
						if (k<names.length) {
							n=ds.addLabel(col,row, names, k, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
							n=ds.addLabel(col,row, myEth.getVarValues(names[k]), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
						} else
							col+=2;
					}
								
					row++;				
				}				
			}
		}
		
		
		row++;
		row++;
		row++;
		
		
		/*
		 * Setup titles for SEA Stats
		 */ 	
		col=0;
		n=ds.addLabel(col,row,"Managed System Name",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"Managed System Serial",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"VIOS",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"SEA",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45); col++;
		n=ds.addLabel(col,row,"Virt Eth",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45); col++;
		n=ds.addLabel(col,row,"Data/Control",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45); col++;
		n=ds.addLabel(col,row,"Priority",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45); col++;
		n=ds.addLabel(col,row,"Active",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45); col++;
		n=ds.addLabel(col,row,"Port VLAN",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45); col++;
		n=ds.addLabel(col,row,"VLANs",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45); col++;
		n=ds.addLabel(col,row,"Switch",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45); col++;
		n=ds.addLabel(col,row,"Hyp Send Failures",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45); col++;
		n=ds.addLabel(col,row,"Hyp Send Fail (LPAR)",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45); col++;
		n=ds.addLabel(col,row,"Hyp Send Fail (VIOS)",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45); col++;
		n=ds.addLabel(col,row,"Hyp Receive Failures",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"Transmit Bufsize",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"Transmit Num Buffers",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"Transmit no buffers",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"Receive MinBufs Tiny",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"Receive MinBufs Small",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"Receive MinBufs Medium",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"Receive MinBufs Large",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"Receive MinBufs Huge",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"Receive MaxBufs Tiny",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"Receive MaxBufs Small",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"Receive MaxBufs Medium",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"Receive MaxBufs Large",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"Receive MaxBufs Huge",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"Receive MaxAlloc Tiny",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"Receive MaxAlloc Small",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"Receive MaxAlloc Medium",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"Receive MaxAlloc Large",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"Receive MaxAlloc Huge",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
			
		row ++;
		
		
		final int normal = VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW;
		final int yellow = VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW|YELLOW;
		final int red = VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW|RED;
		String min[], max[], mAlloc[]; 
		int map;
		boolean hyp_send_fail = false;
		for (i=0; i<managedSystem.length; i++) {
					
			/*
			 * Write virtual slot variables values
			 */
			object = managedSystem[i].getObjects(ENTSTATSEA);
			if (object != null) {
				for (j=0; j<object.length; j++) {
										
					// Get eth mapping object
					eth = object[j].getObjects(ETH);
					if (eth==null)
						continue;
					
					for (k=0; k<eth.length; k++) {
						
						hyp_send_fail=false;
						if (eth[k].getVarValues("hyp_send_failures")!=null)
							if (Float.parseFloat(eth[k].getVarValues("hyp_send_failures")[0])>0)
								hyp_send_fail=true;
					
						col=0; map=normal;
						n=ds.addLabel( col, row, managedSystem[i].getVarValues("name"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
						n=ds.addLabel( col, row, managedSystem[i].getVarValues("serial_num"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
						n=ds.addLabel( col, row, object[j].getVarValues("VIOS"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;					
						n=ds.addLabel( col, row, object[j].getVarValues("SEA"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
						
						n=ds.addLabel( col, row, eth[k].getVarValues("name"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;					
						n=ds.addLabel( col, row, eth[k].getVarValues("control_channel"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
						n=ds.addLabel( col, row, eth[k].getVarValues("trunk"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
						n=ds.addLabel( col, row, eth[k].getVarValues("trunk"), 1, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
						n=ds.addLabel( col, row, eth[k].getVarValues("PVID"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
						n=ds.addLabel( col, row, eth[k].getVarValues("VLAN"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
						n=ds.addLabel( col, row, eth[k].getVarValues("switch"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
						
						if (hyp_send_fail)
							map=yellow;					
						//if (Float.parseFloat(eth[k].getVarValues("hyp_send_failures")[0])>0)
						//	map=yellow;
						
						n=ds.addLabel( col, row, eth[k].getVarValues("hyp_send_failures"), 0, map);  if (n>size[col]) size[col]=n; col++;
						n=ds.addLabel( col, row, eth[k].getVarValues("hyp_send_receiver_fail"), 0, map);  if (n>size[col]) size[col]=n; col++;
						n=ds.addLabel( col, row, eth[k].getVarValues("hyp_send_sender_fail"), 0, map);  if (n>size[col]) size[col]=n; col++;
						map = normal;
						
						if (hyp_send_fail)
							map=yellow;					
						//if (Float.parseFloat(eth[k].getVarValues("hyp_send_failures")[0])>0)
						//	map=yellow;
						
						n=ds.addLabel( col, row, eth[k].getVarValues("hyp_receive_failures"), 0, map);  if (n>size[col]) size[col]=n; col++;
						map = normal;
						
						n=ds.addLabel( col, row, eth[k].getVarValues("bufsize"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
						n=ds.addLabel( col, row, eth[k].getVarValues("numbufs"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
						
						if (hyp_send_fail)
							map=yellow;					
						//if (Float.parseFloat(eth[k].getVarValues("hyp_send_failures")[0])>0)
						//	map=yellow;
						
						n=ds.addLabel( col, row, eth[k].getVarValues("nobufs"), 0, map);  if (n>size[col]) size[col]=n; col++;
						map=normal;
						
						
						min = eth[k].getVarValues("minBufs");
						max = eth[k].getVarValues("maxBufs");
						mAlloc = eth[k].getVarValues("maxAlloc");
						
						for (int x=0; x<5; x++) {
							if (min!=null)
								n=ds.addLabel( col, row, min, x, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
							else n=0;
							if (n>size[col]) 
								size[col]=n; 
							col++;
						}
						
						for (int x=0; x<5; x++) {
							if (max!=null)
								n=ds.addLabel( col, row, max, x, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
							else n=0;
							if (n>size[col]) 
								size[col]=n; 
							col++;
						}
						
						for (int x=0; x<5; x++) {
							if (mAlloc!=null) {
								if (Integer.parseInt(mAlloc[x])==Integer.parseInt(max[x]))
									map = red;
								else if (Integer.parseInt(mAlloc[x])>Integer.parseInt(min[x]))
									map = yellow;
								else
									map = normal;
								n=ds.addLabel( col, row, mAlloc, x, map);
							} else
								n=0;
							if (n>size[col]) 
								size[col]=n; 
							col++;
						}
					
						
						
						row++;		
						
					}
				}				
			}
		}
		

		for (i=0; i<size.length && i<row; i++)
			ds.setColSize(i, size[i]+2);

		return ds;
	}
	
	
	
	private void createSEASheet(WritableSheet sheet)  throws Exception, RowsExceededException, WriteException {
		GenericData object[];
		int row,col;
		int i,j,k;		
		int size[]=new int[49];
		int n;
		String names[];
		int maxVirt = 0;
		GenericData eth[];
		GenericData myEth;
		String vios;
		
			
		row = 0;
		for (i=0; i<size.length; i++)
			size[i] = 4;
		
		
		// Compute max virtual adapters in all SEAs
		for (i=0; i<managedSystem.length; i++) {
			object = managedSystem[i].getObjects(SEA);
			for (j=0; object!=null && j<object.length; j++) {
				names = object[j].getVarValues("virt_adapters");
				if (names.length>maxVirt)
					maxVirt=names.length;
			}
		}
		
		/*
		 * Setup titles 
		 */ 	
		col=0;
		n=addLabel(sheet,col,row,"Managed System Name",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"Managed System Serial",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"VIOS",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"SEA",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45)); col++;
		n=addLabel(sheet,col,row,"ctl_chan",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45)); col++;
		n=addLabel(sheet,col,row,"ctl_chan slot",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45)); col++;
		n=addLabel(sheet,col,row,"gvrp",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45)); col++;
		n=addLabel(sheet,col,row,"ha_mode",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45)); col++;
		n=addLabel(sheet,col,row,"jumbo_frames",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45)); col++;
		n=addLabel(sheet,col,row,"large_receive",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"largesend",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"netaddr",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"accounting",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"lldpsvc",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"qos_mode",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"thread",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"pvid",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"pvid_adapter",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"real_adapter",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"real_adapter slot",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"thread",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		
		for (k=1;k<=maxVirt; k++) {
			n=addLabel(sheet,col,row,"Virt#"+k,formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
			n=addLabel(sheet,col,row,"Virt#"+k+" slot",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		}
			
		
		row ++;
		
		for (i=0; i<managedSystem.length; i++) {
			
			
			/*
			 * Write virtual slot variables values
			 */
			eth = managedSystem[i].getObjects(ETH);
			object = managedSystem[i].getObjects(SEA);
			if (object != null) {
				for (j=0; j<object.length; j++) {
										
					// Get eth mapping object
					vios = object[j].getVarValues("VIOS")[0];
					for (k=0, myEth=null; k<eth.length; k++) {
						myEth = eth[k];
						if (myEth.getVarValues("VIOS")[0].equals(vios))
							break;
					}
					
					col=0;
					n=addLabel(sheet, col, row, managedSystem[i].getVarValues("name"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, managedSystem[i].getVarValues("serial_num"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, object[j].getVarValues("VIOS"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW)); if (n>size[col]) size[col]=n; col++;					
					n=addLabel(sheet, col, row, object[j].getVarValues("SEA"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;				
					n=addLabel(sheet, col, row, object[j].getVarValues("ctl_chan"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					
					n=addLabel(sheet, col, row, myEth.getVarValues(object[j].getVarValues("ctl_chan")[0]), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					
					n=addLabel(sheet, col, row, object[j].getVarValues("gvrp"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, object[j].getVarValues("ha_mode"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, object[j].getVarValues("jumbo_frames"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, object[j].getVarValues("large_receive"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, object[j].getVarValues("largesend"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, object[j].getVarValues("netaddr"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, object[j].getVarValues("accounting"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, object[j].getVarValues("lldpsvc"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, object[j].getVarValues("qos_mode"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, object[j].getVarValues("thread"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, object[j].getVarValues("pvid"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, object[j].getVarValues("pvid_adapter"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, object[j].getVarValues("real_adapter"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					
					n=addLabel(sheet, col, row, myEth.getVarValues(object[j].getVarValues("real_adapter")[0]), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					
					n=addLabel(sheet, col, row, object[j].getVarValues("thread"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;

					names = object[j].getVarValues("virt_adapters");
					for (k=0; k<maxVirt; k++) {
						if (k<names.length) {
							n=addLabel(sheet, col, row, names, k, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
							n=addLabel(sheet, col, row, myEth.getVarValues(names[k]), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
						} else
							col+=2;
					}
								
					row++;				
				}				
			}
		}
		
		row++;
		row++;
		row++;
		
		
		// Compute max virtual adapters in all ETherchannel
		for (i=0; i<managedSystem.length; i++) {
			object = managedSystem[i].getObjects(ETHERCHANNEL);
			for (j=0; object!=null && j<object.length; j++) {
				names = object[j].getVarValues("adapter_names");
				if (names.length>maxVirt)
					maxVirt=names.length;
			}
		}
		
		/*
		 * Setup titles for ETherchannel
		 */ 	
		col=0;
		n=addLabel(sheet,col,row,"Managed System Name",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"Managed System Serial",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"VIOS",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"Etherchannel",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45)); col++;
		n=addLabel(sheet,col,row,"alt_addr",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45)); col++;
		n=addLabel(sheet,col,row,"auto_recovery",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45)); col++;
		n=addLabel(sheet,col,row,"hash_mode",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45)); col++;
		n=addLabel(sheet,col,row,"interval",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45)); col++;
		n=addLabel(sheet,col,row,"mode",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45)); col++;
		n=addLabel(sheet,col,row,"netaddr",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45)); col++;
		n=addLabel(sheet,col,row,"noloss_failover",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"num_retries",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"retry_time",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"use_alt_addr",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"use_jumbo_frame",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"backup_adapter",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"backup_adapter slot",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		
		
		for (k=1;k<=maxVirt; k++) {
			n=addLabel(sheet,col,row,"Adapter#"+k,formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
			n=addLabel(sheet,col,row,"Adapter#"+k+" slot",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		}
			
		
		row ++;
		
		for (i=0; i<managedSystem.length; i++) {
			
			
			/*
			 * Write virtual slot variables values
			 */
			eth = managedSystem[i].getObjects(ETH);
			object = managedSystem[i].getObjects(ETHERCHANNEL);
			if (object != null) {
				for (j=0; j<object.length; j++) {
										
					// Get eth mapping object
					vios = object[j].getVarValues("VIOS")[0];
					for (k=0, myEth=null; k<eth.length; k++) {
						myEth = eth[k];
						if (myEth.getVarValues("VIOS")[0].equals(vios))
							break;
					}
					
					col=0;
					n=addLabel(sheet, col, row, managedSystem[i].getVarValues("name"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, managedSystem[i].getVarValues("serial_num"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, object[j].getVarValues("VIOS"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW)); if (n>size[col]) size[col]=n; col++;					
					n=addLabel(sheet, col, row, object[j].getVarValues("ETHERCHANNEL"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;				
					n=addLabel(sheet, col, row, object[j].getVarValues("alt_addr"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;					
					n=addLabel(sheet, col, row, object[j].getVarValues("auto_recovery"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, object[j].getVarValues("hash_mode"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;					
					n=addLabel(sheet, col, row, object[j].getVarValues("interval"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;					
					n=addLabel(sheet, col, row, object[j].getVarValues("mode"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, object[j].getVarValues("netaddr"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, object[j].getVarValues("noloss_failover"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, object[j].getVarValues("num_retries"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, object[j].getVarValues("retry_time"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, object[j].getVarValues("use_alt_addr"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, object[j].getVarValues("use_jumbo_frame"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, object[j].getVarValues("backup_adapter"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					
					n=addLabel(sheet, col, row, myEth.getVarValues(object[j].getVarValues("backup_adapter")[0]), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					
					names = object[j].getVarValues("adapter_names");
					for (k=0; k<maxVirt; k++) {
						if (k<names.length) {
							n=addLabel(sheet, col, row, names, k, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
							n=addLabel(sheet, col, row, myEth.getVarValues(names[k]), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
						} else
							col+=2;
					}
								
					row++;				
				}				
			}
		}
		
		
		row++;
		row++;
		row++;
		
		
		/*
		 * Setup titles for SEA Stats
		 */ 	
		col=0;
		n=addLabel(sheet,col,row,"Managed System Name",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"Managed System Serial",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"VIOS",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"SEA",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45)); col++;
		n=addLabel(sheet,col,row,"Virt Eth",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45)); col++;
		n=addLabel(sheet,col,row,"Data/Control",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45)); col++;
		n=addLabel(sheet,col,row,"Priority",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45)); col++;
		n=addLabel(sheet,col,row,"Active",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45)); col++;
		n=addLabel(sheet,col,row,"Port VLAN",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45)); col++;
		n=addLabel(sheet,col,row,"VLANs",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45)); col++;
		n=addLabel(sheet,col,row,"Switch",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45)); col++;
		n=addLabel(sheet,col,row,"Hyp Send Failures",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45)); col++;
		n=addLabel(sheet,col,row,"Hyp Send Fail (LPAR)",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45)); col++;
		n=addLabel(sheet,col,row,"Hyp Send Fail (VIOS)",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45)); col++;
		n=addLabel(sheet,col,row,"Hyp Receive Failures",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"Transmit Bufsize",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"Transmit Num Buffers",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"Transmit no buffers",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"Receive MinBufs Tiny",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"Receive MinBufs Small",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"Receive MinBufs Medium",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"Receive MinBufs Large",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"Receive MinBufs Huge",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"Receive MaxBufs Tiny",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"Receive MaxBufs Small",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"Receive MaxBufs Medium",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"Receive MaxBufs Large",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"Receive MaxBufs Huge",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"Receive MaxAlloc Tiny",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"Receive MaxAlloc Small",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"Receive MaxAlloc Medium",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"Receive MaxAlloc Large",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"Receive MaxAlloc Huge",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
			
		row ++;
		
		
		final int normal = VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW;
		final int yellow = VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW|YELLOW;
		final int red = VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW|RED;
		String min[], max[], mAlloc[]; 
		int map;
		boolean hyp_send_fail = false;
		for (i=0; i<managedSystem.length; i++) {
					
			/*
			 * Write virtual slot variables values
			 */
			object = managedSystem[i].getObjects(ENTSTATSEA);
			if (object != null) {
				for (j=0; j<object.length; j++) {
										
					// Get eth mapping object
					eth = object[j].getObjects(ETH);
					if (eth==null)
						continue;
					
					for (k=0; k<eth.length; k++) {
						
						hyp_send_fail=false;
						if (eth[k].getVarValues("hyp_send_failures")!=null)
							if (Float.parseFloat(eth[k].getVarValues("hyp_send_failures")[0])>0)
								hyp_send_fail=true;
					
						col=0; map=normal;
						n=addLabel(sheet, col, row, managedSystem[i].getVarValues("name"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
						n=addLabel(sheet, col, row, managedSystem[i].getVarValues("serial_num"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
						n=addLabel(sheet, col, row, object[j].getVarValues("VIOS"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW)); if (n>size[col]) size[col]=n; col++;					
						n=addLabel(sheet, col, row, object[j].getVarValues("SEA"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
						
						n=addLabel(sheet, col, row, eth[k].getVarValues("name"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;					
						n=addLabel(sheet, col, row, eth[k].getVarValues("control_channel"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
						n=addLabel(sheet, col, row, eth[k].getVarValues("trunk"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
						n=addLabel(sheet, col, row, eth[k].getVarValues("trunk"), 1, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
						n=addLabel(sheet, col, row, eth[k].getVarValues("PVID"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
						n=addLabel(sheet, col, row, eth[k].getVarValues("VLAN"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
						n=addLabel(sheet, col, row, eth[k].getVarValues("switch"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
						
						if (hyp_send_fail)
							map=yellow;					
						//if (Float.parseFloat(eth[k].getVarValues("hyp_send_failures")[0])>0)
						//	map=yellow;
						
						n=addLabel(sheet, col, row, eth[k].getVarValues("hyp_send_failures"), 0, formatLabel(map));  if (n>size[col]) size[col]=n; col++;
						n=addLabel(sheet, col, row, eth[k].getVarValues("hyp_send_receiver_fail"), 0, formatLabel(map));  if (n>size[col]) size[col]=n; col++;
						n=addLabel(sheet, col, row, eth[k].getVarValues("hyp_send_sender_fail"), 0, formatLabel(map));  if (n>size[col]) size[col]=n; col++;
						map = normal;
						
						if (hyp_send_fail)
							map=yellow;					
						//if (Float.parseFloat(eth[k].getVarValues("hyp_send_failures")[0])>0)
						//	map=yellow;
						
						n=addLabel(sheet, col, row, eth[k].getVarValues("hyp_receive_failures"), 0, formatLabel(map));  if (n>size[col]) size[col]=n; col++;
						map = normal;
						
						n=addLabel(sheet, col, row, eth[k].getVarValues("bufsize"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
						n=addLabel(sheet, col, row, eth[k].getVarValues("numbufs"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
						
						if (hyp_send_fail)
							map=yellow;					
						//if (Float.parseFloat(eth[k].getVarValues("hyp_send_failures")[0])>0)
						//	map=yellow;
						
						n=addLabel(sheet, col, row, eth[k].getVarValues("nobufs"), 0, formatLabel(map));  if (n>size[col]) size[col]=n; col++;
						map=normal;
						
						
						min = eth[k].getVarValues("minBufs");
						max = eth[k].getVarValues("maxBufs");
						mAlloc = eth[k].getVarValues("maxAlloc");
						
						for (int x=0; x<5; x++) {
							if (min!=null)
								n=addLabel(sheet, col, row, min, x, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
							else n=0;
							if (n>size[col]) 
								size[col]=n; 
							col++;
						}
						
						for (int x=0; x<5; x++) {
							if (max!=null)
								n=addLabel(sheet, col, row, max, x, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
							else n=0;
							if (n>size[col]) 
								size[col]=n; 
							col++;
						}
						
						for (int x=0; x<5; x++) {
							if (mAlloc!=null) {
								if (Integer.parseInt(mAlloc[x])==Integer.parseInt(max[x]))
									map = red;
								else if (Integer.parseInt(mAlloc[x])>Integer.parseInt(min[x]))
									map = yellow;
								else
									map = normal;
								n=addLabel(sheet, col, row, mAlloc, x, formatLabel(map));
							} else
								n=0;
							if (n>size[col]) 
								size[col]=n; 
							col++;
						}
						
						
						/*
						n=addLabel(sheet, col, row, eth[k].getVarValues("minBufs"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
						n=addLabel(sheet, col, row, eth[k].getVarValues("minBufs"), 1, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
						n=addLabel(sheet, col, row, eth[k].getVarValues("minBufs"), 2, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
						n=addLabel(sheet, col, row, eth[k].getVarValues("minBufs"), 3, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
						n=addLabel(sheet, col, row, eth[k].getVarValues("minBufs"), 4, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
						
						n=addLabel(sheet, col, row, eth[k].getVarValues("maxBufs"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
						n=addLabel(sheet, col, row, eth[k].getVarValues("maxBufs"), 1, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
						n=addLabel(sheet, col, row, eth[k].getVarValues("maxBufs"), 2, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
						n=addLabel(sheet, col, row, eth[k].getVarValues("maxBufs"), 3, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
						n=addLabel(sheet, col, row, eth[k].getVarValues("maxBufs"), 4, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
						
						n=addLabel(sheet, col, row, eth[k].getVarValues("maxAlloc"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
						n=addLabel(sheet, col, row, eth[k].getVarValues("maxAlloc"), 1, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
						n=addLabel(sheet, col, row, eth[k].getVarValues("maxAlloc"), 2, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
						n=addLabel(sheet, col, row, eth[k].getVarValues("maxAlloc"), 3, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
						n=addLabel(sheet, col, row, eth[k].getVarValues("maxAlloc"), 4, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;	
						*/	
						
						
						row++;		
						
					}
				}				
			}
		}
		

		for (i=0; i<size.length; i++)
			sheet.setColumnView(i, size[i]+2);
	}

	
	private void createFCSheetExcel(WritableSheet sheet) {
		DataSheet ds = createFCSheet();
		if (ds!=null)
			ds.createExcelSheet(sheet);
	}
	
	private void createFCSheetHTML(String fileName) {
		DataSheet ds = createFCSheet();
		if (ds!=null) {
			ds.createHTMLSheet(fileName);
			addButton("Physical FC",new File(fileName).getName());
		}
	}
	
	private void createFCSheetCSV(String fileName) {
		DataSheet ds = createFCSheet();
		if (ds!=null) {
			ds.setSeparator(csvSeparator);
			ds.createCSVSheet(fileName);
		}
	}
	
	private DataSheet createFCSheet() {
		DataSheet ds = new DataSheet();
		int row,col;
		int i,j, fc;		
		int size[]=new int[36];
		int n;
		GenericData gd[];
		
			
		row = 0;
		for (i=0; i<size.length; i++)
			size[i] = 4;
		
		
		/*
		 * Setup titles 
		 */ 	
		col=0;
		n=ds.addLabel(col,row,"Managed System Name",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"Managed System Serial",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"VIOS",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"FC",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45); col++;
		n=ds.addLabel(col,row,"DIF_enabled",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45); col++;
		n=ds.addLabel(col,row,"bus_intr_lvl",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45); col++;
		n=ds.addLabel(col,row,"bus_io_addr",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45); col++;
		n=ds.addLabel(col,row,"bus_mem_addr",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45); col++;
		n=ds.addLabel(col,row,"bus_mem_addr2",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45); col++;
		n=ds.addLabel(col,row,"init_link",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"intr_msi_1",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"intr_priority",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"lg_term_dma",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"max_xfer_size",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"num_cmd_elems",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"pref_alpa",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"sw_fc_class",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"tme",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"firmware",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"wwnn",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"wwpn",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"Speed supported",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"Speed running",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"fcid",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"port_type",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"Secs since reset",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"Error frames",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"Dumped frames",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"Invalid CRC",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"Invalid Tx",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"IP no DMA",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"IP no Adapter",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"SCSI no DMA",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"SCSI no Adapter",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"SCSI no Cmd",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
		n=ds.addLabel(col,row,"Effective num_cmds",BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45);  col++;
			
		
		row ++;
		
		for (i=0; i<managedSystem.length; i++) {
			
			
			/*
			 * Write virtual slot variables values
			 */
			gd = managedSystem[i].getObjects(FC);

			if (gd != null) {
				for (j=0; j<gd.length; j++) {
					
					col=0;
					n=ds.addLabel( col, row, managedSystem[i].getVarValues("name"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, managedSystem[i].getVarValues("serial_num"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, gd[j].getVarValues("VIOS"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;					
					n=ds.addLabel( col, row, gd[j].getVarValues("FC"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;				
					n=ds.addLabel( col, row, gd[j].getVarValues("DIF_enabled"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, gd[j].getVarValues("bus_intr_lvl"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, gd[j].getVarValues("bus_io_addr"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, gd[j].getVarValues("bus_mem_addr"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, gd[j].getVarValues("bus_mem_addr2"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, gd[j].getVarValues("init_link"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, gd[j].getVarValues("intr_msi_1"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, gd[j].getVarValues("intr_priority"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, gd[j].getVarValues("lg_term_dma"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, gd[j].getVarValues("max_xfer_size"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, gd[j].getVarValues("num_cmd_elems"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, gd[j].getVarValues("pref_alpa"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, gd[j].getVarValues("sw_fc_class"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, gd[j].getVarValues("tme"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, gd[j].getVarValues("firmware"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, gd[j].getVarValues("wwnn"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, gd[j].getVarValues("wwpn"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, gd[j].getVarValues("speed-supported"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, gd[j].getVarValues("speed-running"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, gd[j].getVarValues("fcid"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, gd[j].getVarValues("port_type"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, gd[j].getVarValues("reset"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, gd[j].getVarValues("errorf"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, gd[j].getVarValues("dumpedf"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, gd[j].getVarValues("invalidcrc"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, gd[j].getVarValues("invalidtx"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, gd[j].getVarValues("ip_nodma"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, gd[j].getVarValues("ip_noadapter"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, gd[j].getVarValues("scsi_nodma"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, gd[j].getVarValues("scsi_noadapter"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, gd[j].getVarValues("scsi_nocmd"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel( col, row, gd[j].getVarValues("e_max_transfer"), 0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);  if (n>size[col]) size[col]=n; col++;
								
					row++;				
				}				
			}
		}
		

		for (i=0; i<size.length; i++)
			ds.setColSize(i, size[i]+2);

		return ds;
	}
	
	
	private void createFCSheet(WritableSheet sheet)  throws Exception, RowsExceededException, WriteException {
		int row,col;
		int i,j, fc;		
		int size[]=new int[36];
		int n;
		GenericData gd[];
		
			
		row = 0;
		for (i=0; i<size.length; i++)
			size[i] = 4;
		
		
		/*
		 * Setup titles 
		 */ 	
		col=0;
		n=addLabel(sheet,col,row,"Managed System Name",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"Managed System Serial",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"VIOS",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"FC",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45)); col++;
		n=addLabel(sheet,col,row,"DIF_enabled",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45)); col++;
		n=addLabel(sheet,col,row,"bus_intr_lvl",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45)); col++;
		n=addLabel(sheet,col,row,"bus_io_addr",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45)); col++;
		n=addLabel(sheet,col,row,"bus_mem_addr",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45)); col++;
		n=addLabel(sheet,col,row,"bus_mem_addr2",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45)); col++;
		n=addLabel(sheet,col,row,"init_link",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"intr_msi_1",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"intr_priority",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"lg_term_dma",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"max_xfer_size",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"num_cmd_elems",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"pref_alpa",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"sw_fc_class",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"tme",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"firmware",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"wwnn",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"wwpn",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"Speed supported",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"Speed running",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"fcid",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"port_type",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"Secs since reset",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"Error frames",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"Dumped frames",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"Invalid CRC",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"Invalid Tx",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"IP no DMA",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"IP no Adapter",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"SCSI no DMA",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"SCSI no Adapter",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"SCSI no Cmd",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
		n=addLabel(sheet,col,row,"Effective num_cmds",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN|DIAG45));  col++;
			
		
		row ++;
		
		for (i=0; i<managedSystem.length; i++) {
			
			
			/*
			 * Write virtual slot variables values
			 */
			gd = managedSystem[i].getObjects(FC);

			if (gd != null) {
				for (j=0; j<gd.length; j++) {
					
					col=0;
					n=addLabel(sheet, col, row, managedSystem[i].getVarValues("name"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, managedSystem[i].getVarValues("serial_num"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, gd[j].getVarValues("VIOS"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW)); if (n>size[col]) size[col]=n; col++;					
					n=addLabel(sheet, col, row, gd[j].getVarValues("FC"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;				
					n=addLabel(sheet, col, row, gd[j].getVarValues("DIF_enabled"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, gd[j].getVarValues("bus_intr_lvl"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, gd[j].getVarValues("bus_io_addr"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, gd[j].getVarValues("bus_mem_addr"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, gd[j].getVarValues("bus_mem_addr2"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, gd[j].getVarValues("init_link"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, gd[j].getVarValues("intr_msi_1"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, gd[j].getVarValues("intr_priority"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, gd[j].getVarValues("lg_term_dma"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, gd[j].getVarValues("max_xfer_size"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, gd[j].getVarValues("num_cmd_elems"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, gd[j].getVarValues("pref_alpa"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, gd[j].getVarValues("sw_fc_class"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, gd[j].getVarValues("tme"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, gd[j].getVarValues("firmware"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, gd[j].getVarValues("wwnn"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, gd[j].getVarValues("wwpn"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, gd[j].getVarValues("speed-supported"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, gd[j].getVarValues("speed-running"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, gd[j].getVarValues("fcid"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, gd[j].getVarValues("port_type"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, gd[j].getVarValues("reset"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, gd[j].getVarValues("errorf"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, gd[j].getVarValues("dumpedf"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, gd[j].getVarValues("invalidcrc"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, gd[j].getVarValues("invalidtx"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, gd[j].getVarValues("ip_nodma"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, gd[j].getVarValues("ip_noadapter"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, gd[j].getVarValues("scsi_nodma"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, gd[j].getVarValues("scsi_noadapter"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, gd[j].getVarValues("scsi_nocmd"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
					n=addLabel(sheet, col, row, gd[j].getVarValues("e_max_transfer"), 0, formatLabel(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));  if (n>size[col]) size[col]=n; col++;
								
					row++;				
				}				
			}
		}
		

		for (i=0; i<size.length; i++)
			sheet.setColumnView(i, size[i]+2);
	}
	
	
	private void createVFCSheet(WritableSheet sheet)  throws RowsExceededException, WriteException {
		GenericData object[];
		GenericData vfcObject[];
		int row;
		int i,j,k;
		String s[];
		String viosName;	// name of VIOS providing physical adapter
		String viosSlot;	// vfchost slot
		boolean isVios;		// true if current lpar is a VIOS
		int size[]=new int[10];
		int n;
	
		row = 0;
		for (i=0; i<size.length; i++)
			size[i] = 0;

		for (i=0; i<managedSystem.length; i++) {
			
			/*
			 * Show start of system
			 */
			s = managedSystem[i].getVarValues("name");
			sheet.mergeCells(0, row, 9, row);
			addLabel(sheet,0,row,s[0],formatLabel(BOLD|CENTRE|VCENTRE|GREEN));

			row++;
			
			
			/*
			 * Setup titles for virtual slots
			 */ 	
			n=addLabel(sheet,0,row,"LPAR name",formatLabel(BOLD|CENTRE|B_ALL_MED));
			if (n>size[0]) size[0]=n;
			n=addLabel(sheet,1,row,"Slot",formatLabel(BOLD|CENTRE|B_ALL_MED));
			if (n>size[1]) size[1]=n;
			n=addLabel(sheet,2,row,"State",formatLabel(BOLD|CENTRE|B_ALL_MED));
			if (n>size[2]) size[2]=n;
			n=addLabel(sheet,3,row,"Required",formatLabel(BOLD|CENTRE|B_ALL_MED));
			if (n>size[3]) size[3]=n;
			n=addLabel(sheet,4,row,"Type",formatLabel(BOLD|CENTRE|B_ALL_MED));
			if (n>size[4]) size[4]=n;
			n=addLabel(sheet,5,row,"Remote LPAR",formatLabel(BOLD|CENTRE|B_ALL_MED));
			if (n>size[5]) size[5]=n;
			n=addLabel(sheet,6,row,"Remote Slot",formatLabel(BOLD|CENTRE|B_ALL_MED));
			if (n>size[6]) size[6]=n;		
			n=addLabel(sheet,7,row,"WWPN #1",formatLabel(BOLD|CENTRE|B_ALL_MED));
			if (n>size[7]) size[7]=n;
			n=addLabel(sheet,8,row,"WWPN #2",formatLabel(BOLD|CENTRE|B_ALL_MED));
			if (n>size[8]) size[8]=n;
			n=addLabel(sheet,9,row,"Physical FC Slot",formatLabel(BOLD|CENTRE|B_ALL_MED));
			if (n>size[9]) size[9]=n;
			
			
			row ++;
			
			/*
			 * Write virtual slot variables values
			 */
			viosName = viosSlot = null;
			object = managedSystem[i].getObjects(VFC);
			if (object != null) {
				for (j=0; j<object.length; j++) {
					
					s = object[j].getVarValues("lpar_name");
					n=addLabel(sheet, 0, row, s,0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					if (n>size[0]) size[0]=n;
					if (s!=null)
						viosName = s[0];
									
					s = object[j].getVarValues("slot_num");
					addNumber(sheet, 1, row, s,0, formatInt(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					if (s!=null)
						viosSlot = s[0];
					
					s = object[j].getVarValues("state");
					if (s!=null) {
						if (s[0].equals("1"))
							addLabel(sheet, 2, row, "On", formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
						else
							addLabel(sheet, 2, row, "Off", formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					}
					
					s = object[j].getVarValues("is_required");
					if (s!=null) {
						if (s[0].equals("1"))
							addLabel(sheet, 3, row, "True", formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
						else
							addLabel(sheet, 3, row, "False", formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					}
					
					s = object[j].getVarValues("adapter_type");
					n=addLabel(sheet, 4, row, s,0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					if (n>size[4]) size[4]=n;
					if (s!=null && s[0].equals("server"))
						isVios = true;
					else
						isVios = false;
					
					s = object[j].getVarValues("remote_lpar_name");
					n=addLabel(sheet, 5, row, s,0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					if (n>size[5]) size[5]=n;
					if (!isVios && s!=null)
						viosName = s[0];
					
					s = object[j].getVarValues("remote_slot_num");
					if (s!=null && s[0].equals("any"))
						addLabel(sheet, 6, row, s,0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					else
						addNumber(sheet, 6, row, s,0, formatInt(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					if (!isVios && s!=null)
						viosSlot = s[0];
					
					s = object[j].getVarValues("wwpns");
					n=addLabel(sheet, 7, row, s,0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					if (n>size[7]) size[7]=n;
					n=addLabel(sheet, 8, row, s,1, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					if (n>size[8]) size[8]=n;
					
					vfcObject = managedSystem[i].getObjects(VFCMAP);	
					if (vfcObject != null) {
						for (k=0; k<vfcObject.length; k++) {
							s = vfcObject[k].getVarValues(viosName+"@"+viosSlot);
							if (s!=null)
								break;
						}
						n=addLabel(sheet, 9, row, s,0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
						if (n>size[9]) size[9]=n;
					}
								
					row++;				
				}				
			}
			
			
			row +=2;
		}
		
		for (i=0; i<size.length; i++)
			sheet.setColumnView(i, size[i]+2);
	}
	
	
	
	
	private void createVFCSheetRowBasedExcel(WritableSheet sheet) {
		DataSheet ds = createVFCSheetRowBased();
		if (ds!=null)
			ds.createExcelSheet(sheet);
	}
	
	private void createVFCSheetRowBasedHTML(String fileName) {
		DataSheet ds = createVFCSheetRowBased();
		if (ds!=null) {
			ds.createHTMLSheet(fileName);
			addButton("Virtual Fibre",new File(fileName).getName());
		}
	}
	
	private void createVFCSheetRowBasedCSV(String fileName) {
		DataSheet ds = createVFCSheetRowBased();
		if (ds!=null) {
			ds.setSeparator(csvSeparator);
			ds.createCSVSheet(fileName);
		}
	}
	
	private DataSheet createVFCSheetRowBased()  {
		DataSheet ds = new DataSheet();
		GenericData object[];
		GenericData vfcObject[];
		int row, col;
		int i,j,k;
		String s[];
		String viosName;	// name of VIOS providing physical adapter
		String viosSlot;	// vfchost slot
		boolean isVios;		// true if current lpar is a VIOS
		int size[]=new int[12];
		int n;
	
		row = 0;
		col = 0;
		for (i=0; i<size.length; i++)
			size[i] = 0;
		
		
		/*
		 * Setup titles for virtual slots
		 */ 	
		n=ds.addLabel(col,row,"LPAR name",BOLD|CENTRE|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n=ds.addLabel(col,row,"Slot",BOLD|CENTRE|B_ALL_MED|GREEN);if (n>size[col]) size[col]=n; col++;
		n=ds.addLabel(col,row,"State",BOLD|CENTRE|B_ALL_MED|GREEN);if (n>size[col]) size[col]=n; col++;
		n=ds.addLabel(col,row,"Required",BOLD|CENTRE|B_ALL_MED|GREEN);if (n>size[col]) size[col]=n; col++;
		n=ds.addLabel(col,row,"Type",BOLD|CENTRE|B_ALL_MED|GREEN);if (n>size[col]) size[col]=n; col++;
		n=ds.addLabel(col,row,"Remote LPAR",BOLD|CENTRE|B_ALL_MED|GREEN);if (n>size[col]) size[col]=n; col++;
		n=ds.addLabel(col,row,"Remote Slot",BOLD|CENTRE|B_ALL_MED|GREEN);if (n>size[col]) size[col]=n; col++;
		n=ds.addLabel(col,row,"WWPN #1",BOLD|CENTRE|B_ALL_MED|GREEN);if (n>size[col]) size[col]=n; col++;
		n=ds.addLabel(col,row,"WWPN #2",BOLD|CENTRE|B_ALL_MED|GREEN);if (n>size[col]) size[col]=n; col++;
		n=ds.addLabel(col,row,"Physical FC Slot",BOLD|CENTRE|B_ALL_MED|GREEN);if (n>size[col]) size[col]=n; col++;
		n=ds.addLabel(col,row,"Managed System Name",BOLD|CENTRE|B_ALL_MED|GREEN);if (n>size[col]) size[col]=n; col++;
		n=ds.addLabel(col,row,"Managed System Serial",BOLD|CENTRE|B_ALL_MED|GREEN);if (n>size[col]) size[col]=n; col++;
		
		
		row ++;

		for (i=0; i<managedSystem.length; i++) {
			
			/*
			 * Write virtual slot variables values
			 */
			viosName = viosSlot = null;
			object = managedSystem[i].getObjects(VFC);
			if (object != null) {
				for (j=0; j<object.length; j++) {
					
					col = 0;
					
					s = object[j].getVarValues("lpar_name");
					n=ds.addLabel(col,row, s,0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
					if (s!=null)
						viosName = s[0];
									
					s = object[j].getVarValues("slot_num");
					ds.addInteger(col, row, s,0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
					n=5; if (n>size[col]) size[col]=n; col++;
					if (s!=null)
						viosSlot = s[0];
					
					s = object[j].getVarValues("state");
					if (s!=null) {
						if (s[0].equals("1"))
							n=ds.addLabel(col,row, "On", B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
						else
							n=ds.addLabel(col,row, "Off", B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
						if (n>size[col]) size[col]=n;
					}
					col++;
					
					s = object[j].getVarValues("is_required");
					if (s!=null) {
						if (s[0].equals("1"))
							n=ds.addLabel(col,row, "True", B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
						else
							n=ds.addLabel(col,row, "False", B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
						if (n>size[col]) size[col]=n;
					}
					col++;
					
					s = object[j].getVarValues("adapter_type");
					n=ds.addLabel(col,row, s,0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
					if (s!=null && s[0].equals("server"))
						isVios = true;
					else
						isVios = false;
					
					s = object[j].getVarValues("remote_lpar_name");
					n=ds.addLabel(col,row, s,0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
					if (!isVios && s!=null)
						viosName = s[0];
					
					s = object[j].getVarValues("remote_slot_num");
					if (s!=null && s[0].equals("any"))
						n=ds.addLabel(col, row, s,0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
					else {
						ds.addInteger(col, row, s,0, VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
						n=5;
					}
					if (n>size[col]) size[col]=n; col++;
					if (!isVios && s!=null)
						viosSlot = s[0];
					
					s = object[j].getVarValues("wwpns");
					n=ds.addLabel(col,row, s,0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel(col,row, s,1, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
					
					vfcObject = managedSystem[i].getObjects(VFCMAP);	
					if (vfcObject != null) {
						for (k=0; k<vfcObject.length; k++) {
							s = vfcObject[k].getVarValues(viosName+"@"+viosSlot);
							if (s!=null)
								break;
						}
						n=ds.addLabel(col,row, s,0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
					} else
						col++;
					
					n=ds.addLabel(col,row, managedSystem[i].getVarValues("name"),0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
					n=ds.addLabel(col,row, managedSystem[i].getVarValues("serial_num"),0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
								
					row++;				
				}				
			}
		}
		
		for (i=0; i<size.length; i++)
			ds.setColSize(i, size[i]+2);

		return ds;
	}
	
	
	private void createVFCSheetRowBased(WritableSheet sheet)  throws RowsExceededException, WriteException {
		GenericData object[];
		GenericData vfcObject[];
		int row;
		int i,j,k;
		String s[];
		String viosName;	// name of VIOS providing physical adapter
		String viosSlot;	// vfchost slot
		boolean isVios;		// true if current lpar is a VIOS
		int size[]=new int[12];
		int n;
	
		row = 0;
		for (i=0; i<size.length; i++)
			size[i] = 0;
		
		
		/*
		 * Setup titles for virtual slots
		 */ 	
		n=addLabel(sheet,0,row,"LPAR name",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN));
		if (n>size[0]) size[0]=n;
		n=addLabel(sheet,1,row,"Slot",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN));
		if (n>size[1]) size[1]=n;
		n=addLabel(sheet,2,row,"State",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN));
		if (n>size[2]) size[2]=n;
		n=addLabel(sheet,3,row,"Required",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN));
		if (n>size[3]) size[3]=n;
		n=addLabel(sheet,4,row,"Type",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN));
		if (n>size[4]) size[4]=n;
		n=addLabel(sheet,5,row,"Remote LPAR",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN));
		if (n>size[5]) size[5]=n;
		n=addLabel(sheet,6,row,"Remote Slot",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN));
		if (n>size[6]) size[6]=n;		
		n=addLabel(sheet,7,row,"WWPN #1",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN));
		if (n>size[7]) size[7]=n;
		n=addLabel(sheet,8,row,"WWPN #2",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN));
		if (n>size[8]) size[8]=n;
		n=addLabel(sheet,9,row,"Physical FC Slot",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN));
		if (n>size[9]) size[9]=n;
		n=addLabel(sheet,10,row,"Managed System Name",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN));
		if (n>size[10]) size[10]=n;
		n=addLabel(sheet,11,row,"Managed System Serial",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN));
		if (n>size[11]) size[11]=n;
		
		
		row ++;

		for (i=0; i<managedSystem.length; i++) {
			
			/*
			 * Write virtual slot variables values
			 */
			viosName = viosSlot = null;
			object = managedSystem[i].getObjects(VFC);
			if (object != null) {
				for (j=0; j<object.length; j++) {
					
					s = object[j].getVarValues("lpar_name");
					n=addLabel(sheet, 0, row, s,0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					if (n>size[0]) size[0]=n;
					if (s!=null)
						viosName = s[0];
									
					s = object[j].getVarValues("slot_num");
					addNumber(sheet, 1, row, s,0, formatInt(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					if (s!=null)
						viosSlot = s[0];
					
					s = object[j].getVarValues("state");
					if (s!=null) {
						if (s[0].equals("1"))
							addLabel(sheet, 2, row, "On", formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
						else
							addLabel(sheet, 2, row, "Off", formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					}
					
					s = object[j].getVarValues("is_required");
					if (s!=null) {
						if (s[0].equals("1"))
							addLabel(sheet, 3, row, "True", formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
						else
							addLabel(sheet, 3, row, "False", formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					}
					
					s = object[j].getVarValues("adapter_type");
					n=addLabel(sheet, 4, row, s,0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					if (n>size[4]) size[4]=n;
					if (s!=null && s[0].equals("server"))
						isVios = true;
					else
						isVios = false;
					
					s = object[j].getVarValues("remote_lpar_name");
					n=addLabel(sheet, 5, row, s,0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					if (n>size[5]) size[5]=n;
					if (!isVios && s!=null)
						viosName = s[0];
					
					s = object[j].getVarValues("remote_slot_num");
					if (s!=null && s[0].equals("any"))
						addLabel(sheet, 6, row, s,0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					else
						addNumber(sheet, 6, row, s,0, formatInt(VCENTRE|B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					if (!isVios && s!=null)
						viosSlot = s[0];
					
					s = object[j].getVarValues("wwpns");
					n=addLabel(sheet, 7, row, s,0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					if (n>size[7]) size[7]=n;
					n=addLabel(sheet, 8, row, s,1, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					if (n>size[8]) size[8]=n;
					
					vfcObject = managedSystem[i].getObjects(VFCMAP);	
					if (vfcObject != null) {
						for (k=0; k<vfcObject.length; k++) {
							s = vfcObject[k].getVarValues(viosName+"@"+viosSlot);
							if (s!=null)
								break;
						}
						n=addLabel(sheet, 9, row, s,0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
						if (n>size[9]) size[9]=n;
					}
					
					n = addLabel(sheet, 10, row, managedSystem[i].getVarValues("name"),0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					if (n>size[10]) size[10]=n;
					n = addLabel(sheet, 11, row, managedSystem[i].getVarValues("serial_num"),0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
					if (n>size[11]) size[11]=n;
								
					row++;				
				}				
			}
		}
		
		for (i=0; i<size.length; i++)
			sheet.setColumnView(i, size[i]+2);
	}
	
	
	private void createGlobalMemSheet(WritableSheet sheet)  throws RowsExceededException, WriteException {
		String systemName;
		GenericData object[];
		int row;
		String names[];
		int i,j,k,x;
		String s[];
		Label label;
		String v;

		
		WritableFont arial10font = new WritableFont(WritableFont.ARIAL, 10);
		WritableFont arial10boldfont = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD, true);
		WritableCellFormat arial10format = new WritableCellFormat (arial10font);
		WritableCellFormat arial10boldformat = new WritableCellFormat (arial10boldfont);
		

		row = 0;
		
		// Merge variable names
		names=null;
		for (i=0; i<managedSystem.length; i++) {		
			object = managedSystem[i].getObjects(MEM_LPAR);
			if (object==null)
				continue;			
			for (j=0; j<object.length; j++)  {
				s = object[j].getVarNames();
				if (names==null)
					names=s;
				else
					names = mergeList(names, s);
			}
		}
			
		row=0;
		for (i=0; i<managedSystem.length; i++) {
			systemName = managedSystem[i].getVarValues("name")[0];
			
			label = new Label(0,row,systemName+" MEM",arial10boldformat);
			sheet.addCell(label);
			row++;
			
			object = managedSystem[i].getObjects(MEM_LPAR);
			if (object==null) {
				row+=2;
				continue;
			}
			
			// Print variable names				
			for (j=0; j<names.length; j++) {
				label = new Label(j,row,names[j],arial10boldformat);
				sheet.addCell(label);					
			}
			row++;
			
			for (j=0; j<object.length; j++) {
				for (k=0; k<names.length; k++) {
					s = object[j].getVarValues(names[k]);
					if (s==null)
						continue;
					
					v=null;
					for (x=0; x<s.length; x++) {	
						if (v==null)
							v=s[x];
						else
							v=v+","+s[x];
					}
				
					label = new Label(k,row,v,arial10format);
					sheet.addCell(label);
				}
				row++;
			}
			
			object = managedSystem[i].getObjects(MEM_POOL);
			if (object!=null) {
				row++;
				
				String names2[]=null;
				for (j=0; j<object.length; j++)  {
					s = object[j].getVarNames();
					if (names2==null)
						names2=s;
					else
						names2 = mergeList(names2, s);
				}
				
				for (j=0; j<names2.length; j++) {
					label = new Label(j,row,names2[j],arial10boldformat);
					sheet.addCell(label);					
				}
				row++;
				
				for (j=0; j<object.length; j++) {
					for (k=0; k<names2.length; k++) {
						s = object[j].getVarValues(names2[k]);
						if (s==null)
							continue;
						
						v=null;
						for (x=0; x<s.length; x++) {	
							if (v==null)
								v=s[x];
							else
								v=v+","+s[x];
						}
					
						label = new Label(k,row,v,arial10format);
						sheet.addCell(label);
					}
					row++;
				}					
			}
			
			row+=2;							
		}		
		
	}
	
	
	
	
	
	
	
	private void createGlobalCPUSheet(WritableSheet sheet)  throws RowsExceededException, WriteException {
		String systemName;
		GenericData object[];
		int row;
		String names[];
		int i,j,k,x;
		String s[];
		Label label;
		String v;
		
		WritableFont arial10font = new WritableFont(WritableFont.ARIAL, 10);
		WritableFont arial10boldfont = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD, true);
		WritableCellFormat arial10format = new WritableCellFormat (arial10font);
		WritableCellFormat arial10boldformat = new WritableCellFormat (arial10boldfont);
		
		WritableCellFormat centre = new WritableCellFormat (arial10boldfont);
		centre.setAlignment(Alignment.CENTRE);
		centre.setVerticalAlignment(VerticalAlignment.CENTRE);
		centre.setBackground(Colour.LIGHT_GREEN);
		
		WritableCellFormat title = new WritableCellFormat (arial10boldfont);
		title.setWrap(true);
		

		row = 0;
		
		// Merge variable names
		names=null;
		for (i=0; i<managedSystem.length; i++) {		
			object = managedSystem[i].getObjects(PROC_LPAR);
			if (object==null)
				continue;			
			for (j=0; j<object.length; j++)  {
				s = object[j].getVarNames();
				if (names==null)
					names=s;
				else
					names = mergeList(names, s);
			}
		}
		

		
		
		int columnSize[] = new int[names.length];
		for (i=0; i<columnSize.length; i++)
			columnSize[i]=0;
		
			
		row=0;
		for (i=0; i<managedSystem.length; i++) {
			systemName = managedSystem[i].getVarValues("name")[0];

			sheet.mergeCells(0, row, names.length-1, row);
			label = new Label(0,row,systemName+" CPU",centre);
			sheet.addCell(label);
			row++;
			
			object = managedSystem[i].getObjects(PROC_LPAR);
			if (object==null) {
				row+=2;
				continue;
			}
			
			// Print variable names				
			for (j=0; j<names.length; j++) {
				label = new Label(j,row,names[j].replace('_', ' '),title);
				sheet.addCell(label);					
			}
			row++;
			
			for (j=0; j<object.length; j++) {
				for (k=0; k<names.length; k++) {
					s = object[j].getVarValues(names[k]);
					if (s==null)
						continue;
					
					v=null;
					for (x=0; x<s.length; x++) {	
						if (v==null)
							v=s[x];
						else
							v=v+","+s[x];
					}
				
					label = new Label(k,row,v,arial10format);
					sheet.addCell(label);
					
					if (v.length()>columnSize[k])
						columnSize[k]=v.length();
				}
				row++;
			}
			
			object = managedSystem[i].getObjects(PROC_POOL);
			if (object!=null) {
				row++;
				
				String names2[]=null;
				for (j=0; j<object.length; j++)  {
					s = object[j].getVarNames();
					if (names2==null)
						names2=s;
					else
						names2 = mergeList(names2, s);
				}

				
				for (j=0; j<names2.length; j++) {
					label = new Label(j,row,names2[j].replace('_', ' '),title);
					sheet.addCell(label);					
				}
				row++;
				
				for (j=0; j<object.length; j++) {
					for (k=0; k<names2.length; k++) {
						s = object[j].getVarValues(names2[k]);
						if (s==null)
							continue;
						
						v=null;
						for (x=0; x<s.length; x++) {	
							if (v==null)
								v=s[x];
							else
								v=v+","+s[x];
						}
					
						label = new Label(k,row,v,arial10format);
						sheet.addCell(label);
						
						if (v.length()>columnSize[k])
							columnSize[k]=v.length();
					}
					row++;
				}					
			}
			
			row+=2;						
		}
		
		for (i=0; i<columnSize.length; i++)
			sheet.setColumnView(i, 12);

	}
	
	
	
	private void createCPUSheetRowBasedExcel(WritableSheet sheet) {
		DataSheet ds = createCPUSheetRowBased();
		if (ds!=null)
			ds.createExcelSheet(sheet);
	}
	
	private void createCPUSheetRowBasedHTML(String fileName) {
		DataSheet ds = createCPUSheetRowBased();
		if (ds!=null) {
			ds.createHTMLSheet(fileName);
			addButton("LPAR CPU",new File(fileName).getName());
		}
	}
	
	private void createCPUSheetRowBasedCSV(String fileName) {
		DataSheet ds = createCPUSheetRowBased();
		if (ds!=null) {
			ds.setSeparator(csvSeparator);
			ds.createCSVSheet(fileName);
		}
	}
	
	private DataSheet createCPUSheetRowBased() {
		DataSheet ds = new DataSheet();
		GenericData lpar[], pool[];
		int row, col;
		int i,j,k;
		String s[];
		String str;
		double vpe;
		
		String currentProfileName = null;
		GenericData currentProfileData = null;
		int map;
		double active, desired;
		boolean poweredOff;
		String profMode;
		boolean profModeChange=false;
		
		int size[]=new int[17];
		int n;
		for (i=0; i<size.length; i++)
			size[i] = 0;
		
			
		row = 0;
		col = 0;
		
		
		/*
		 * Setup titles
		 */ 
		n = ds.addLabel(col,row,"Name",BOLD|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Status",BOLD|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Mode",BOLD|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Min CPU",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Curr CPU",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Max CPU",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Min Ent",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Curr Ent",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Max Ent",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Weight",BOLD|CENTRE|VCENTRE|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Sharing Mode",BOLD|CENTRE|VCENTRE|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		
		n = ds.addLabel(col,row,"Pool Name",BOLD|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Pool Resv",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Pool Max",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		
		n = ds.addLabel(col,row,"VP:E %",BOLD|CENTRE|VCENTRE|B_ALL_MED|GREEN);		 if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Managed System Name",BOLD|CENTRE|VCENTRE|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Managed System Serial",BOLD|CENTRE|VCENTRE|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		
		row++;
		

		for (i=0; i<managedSystem.length; i++) {
						
			/*
			 * Write variables
			 */
			lpar = managedSystem[i].getObjects(PROC_LPAR);
			if (lpar==null) {
				continue;
			}
			
			map = B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW;
			
			for (j=0; j<lpar.length; j++) {
				
				col = 0;
				
				s = lpar[j].getVarValues("lpar_name");
				n = ds.addLabel(col, row, s, 0, map); if (n>size[col]) size[col]=n; col++;
				
				currentProfileName = getActiveProfileName(i,s[0]);
				currentProfileData = getProfileData(i,s[0], currentProfileName);
				
				poweredOff=false;
				s = lpar[j].getVarValues("run_procs");
				if (s[0].equals("0") || s[0].equals("null")) {
					str = "off";
					poweredOff=true;
				} else {
					str = "On";
				}
				n = ds.addLabel(col, row, str, map); if (n>size[col]) size[col]=n; col++;
				
				
				s = lpar[j].getVarValues("curr_proc_mode");
				if (currentProfileData!=null && !poweredOff) {
					if (currentProfileData.getVarValues("all_resources")[0].equals("1"))
						profMode="ded";
					else
						profMode = currentProfileData.getVarValues("proc_mode")[0];
					if (!profMode.equals(s[0])) {
						map = map | YELLOW;
						profModeChange=true;
					}
				}				
				n = ds.addLabel(col, row, lpar[j].getVarValues("curr_proc_mode"),0, map); if (n>size[col]) size[col]=n; col++;
				map = B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW;
				
				ds.addInteger(col, row, lpar[j].getVarValues("curr_min_procs"),0, map); n=3; if (n>size[col]) size[col]=n; col++;
				
				if (!profModeChange && currentProfileData!=null && !poweredOff) {
					active = Double.parseDouble(lpar[j].getVarValues("run_procs")[0]);
					s = currentProfileData.getVarValues("desired_procs");
					if (s!=null) {
						desired = Double.parseDouble(s[0]);
						if (active!=desired)
							map = map | YELLOW;
					}
				}
				ds.addInteger(col, row, lpar[j].getVarValues("run_procs"),0, map); n=3; if (n>size[col]) size[col]=n; col++;
				map = B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW;
				
				ds.addInteger(col, row, lpar[j].getVarValues("curr_max_procs"),0, map); n=3; if (n>size[col]) size[col]=n; col++;
				ds.addFloat(col, row, lpar[j].getVarValues("curr_min_proc_units"),0, map); n=6; if (n>size[col]) size[col]=n; col++;
				
				s = lpar[j].getVarValues("run_proc_units");
				if (!profModeChange && currentProfileData!=null && s!=null && !poweredOff) {
					active = Double.parseDouble(lpar[j].getVarValues("run_proc_units")[0]);
					desired = Double.parseDouble(currentProfileData.getVarValues("desired_proc_units")[0]); 
					if (active!=desired) 
						map = map | YELLOW;	
				}
				ds.addFloat(col, row, lpar[j].getVarValues("run_proc_units"),0, map); n=6; if (n>size[col]) size[col]=n; col++;
				map = B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW;
				
				ds.addFloat(col, row, lpar[j].getVarValues("curr_max_proc_units"),0, map); n=6; if (n>size[col]) size[col]=n; col++;
				ds.addInteger(col, row, lpar[j].getVarValues("run_uncap_weight"),0, map); n=3; if (n>size[col]) size[col]=n; col++;
				n = ds.addLabel(col, row, lpar[j].getVarValues("curr_sharing_mode"),0, map); if (n>size[col]) size[col]=n; col++;		
				
				s = lpar[j].getVarValues("curr_shared_proc_pool_name");
				
				if (s!=null) {
					pool = managedSystem[i].getObjects(PROC_POOL);
					for (k=0; k<pool.length; k++)
						if (pool[k].getVarValues("name")[0].equals(s[0]))
							break;
					
					if (k!=pool.length) {
						n = ds.addLabel(col, row, s, 0, map); if (n>size[col]) size[col]=n; col++;
						ds.addFloat(col, row, pool[k].getVarValues("curr_reserved_pool_proc_units"),0, map); n=6; if (n>size[col]) size[col]=n; col++;
						ds.addInteger(col, row, pool[k].getVarValues("max_pool_proc_units"),0, map); n=3; if (n>size[col]) size[col]=n; col++;
					} else
						col = col + 3;
				} else
					col = col + 3;
				
				vpe = 0;
				if (poweredOff)
					str = "Off";
				else if (lpar[j].getVarValues("curr_proc_mode")[0].equals("ded"))
					str = "Dedicated";
				else if (lpar[j].getVarValues("run_procs")[0].equals("1"))
					str = "MicroLPAR";
				else  {
					vpe = 100f * 
							Double.parseDouble(lpar[j].getVarValues("run_procs")[0]) / 
							Double.parseDouble(lpar[j].getVarValues("run_proc_units")[0]); 
				}
				if (vpe>0)
					ds.addFloat(col, row, vpe, map);
				else
					ds.addLabel(col, row, str, map);
				n=8; if (n>size[col]) size[col]=n; col++;
				
				// O113 --> =IF(C113="ded","Dedicated",IF(E113=0,"off",IF(E113>1,ROUND(E113/H113*100,0),"MicroLPAR")))
				//f = new Formula(14, row, "IF(C"+(row+1)+"=\"ded\",\"Dedicated\",IF(E"+(row+1)+"=0,\"Off\",IF(E"+(row+1)+">1,ROUND(E"+(row+1)+"/H"+(row+1)+"*100,0),\"MicroLPAR\")))",formatFloat(map));
				//sheet.addCell(f); 
				
				n = ds.addLabel(col, row, managedSystem[i].getVarValues("name"),0, map); if (n>size[col]) size[col]=n; col++;
				n = ds.addLabel(col, row, managedSystem[i].getVarValues("serial_num"),0, map); if (n>size[col]) size[col]=n; col++;
							
				row++;					
			}		
		}

		for (i=0; i<size.length; i++)
			ds.setColSize(i, size[i]+2);

		return ds;
	}
	
	
	private void createCPUSheetRowBased(WritableSheet sheet)  throws RowsExceededException, WriteException {
		GenericData lpar[], pool[], confLpar[], profile[];
		int row;
		int i,j,k;
		String s[];
		Label label;

		int lparNameSize=0;
		int poolNameSize=0;
		
		String currentProfileName = null;
		GenericData currentProfileData = null;
		int map;
		double active, desired;
		boolean poweredOff;
		
		
		Formula f;
		
		row = 0;
		
		
		
		/*
		 * Setup titles
		 */ 
		addLabel(sheet,0,row,"Name",formatLabel(BOLD|B_ALL_MED|GREEN));
		addLabel(sheet,1,row,"Status",formatLabel(BOLD|B_ALL_MED|GREEN));
		addLabel(sheet,2,row,"Mode",formatLabel(BOLD|B_ALL_MED|GREEN));		
		addLabel(sheet,3,row,"Min CPU",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN));
		addLabel(sheet,4,row,"Curr CPU",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN));
		addLabel(sheet,5,row,"Max CPU",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN));
		addLabel(sheet,6,row,"Min Ent",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN));
		addLabel(sheet,7,row,"Curr Ent",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN));
		addLabel(sheet,8,row,"Max Ent",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN));
		addLabel(sheet,9,row,"Weight",formatLabel(BOLD|CENTRE|VCENTRE|B_ALL_MED|GREEN));
		addLabel(sheet,10,row,"Sharing Mode",formatLabel(BOLD|CENTRE|VCENTRE|B_ALL_MED|GREEN));
		
		addLabel(sheet,11,row,"Pool Name",formatLabel(BOLD|B_ALL_MED|GREEN));
		addLabel(sheet,12,row,"Pool Resv",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN));
		addLabel(sheet,13,row,"Pool Max",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN));
		
		addLabel(sheet,14,row,"E:VP %",formatLabel(BOLD|CENTRE|VCENTRE|B_ALL_MED|GREEN));		
		addLabel(sheet,15,row,"Managed System Name",formatLabel(BOLD|CENTRE|VCENTRE|B_ALL_MED|GREEN));
		addLabel(sheet,16,row,"Managed System Serial",formatLabel(BOLD|CENTRE|VCENTRE|B_ALL_MED|GREEN));
		
		row++;
		

		for (i=0; i<managedSystem.length; i++) {
						
			/*
			 * Write variables
			 */
			lpar = managedSystem[i].getObjects(PROC_LPAR);
			if (lpar==null) {
				continue;
			}
			
			map = B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW;
			
			for (j=0; j<lpar.length; j++) {
				
				s = lpar[j].getVarValues("lpar_name");
				addLabel(sheet, 0, row, s, 0, formatLabel(map));
				if (s[0].length()>lparNameSize)
					lparNameSize=s[0].length();
				
				currentProfileName = getActiveProfileName(i,s[0]);
				currentProfileData = getProfileData(i,s[0], currentProfileName);
				
				poweredOff=false;
				if (lpar[j].getVarValues("run_procs")[0].equals("0")) {
					label = new Label(1, row, "Off", formatLabel(map));
					poweredOff=true;
				} else {
					label = new Label(1, row, "On", formatLabel(map));
				}
				sheet.addCell(label);
				
				addLabel(sheet, 2, row, lpar[j].getVarValues("curr_proc_mode"),0, formatLabel(map));
				addNumber(sheet, 3, row, lpar[j].getVarValues("curr_min_procs"),0, formatInt(map));
				
				if (currentProfileData!=null && !poweredOff) {
					active = Double.parseDouble(lpar[j].getVarValues("run_procs")[0]);
					s = currentProfileData.getVarValues("desired_procs");
					if (s!=null) {
						desired = Double.parseDouble(s[0]);
						if (active!=desired)
							map = map | YELLOW;
					}
				}
				addNumber(sheet, 4, row, lpar[j].getVarValues("run_procs"),0, formatInt(map));
				map = B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW;
				
				addNumber(sheet, 5, row, lpar[j].getVarValues("curr_max_procs"),0, formatInt(map));
				addNumber(sheet, 6, row, lpar[j].getVarValues("curr_min_proc_units"),0, formatFloat(map));
				
				s = lpar[j].getVarValues("run_proc_units");
				if (currentProfileData!=null && s!=null && !poweredOff) {
					active = Double.parseDouble(lpar[j].getVarValues("run_proc_units")[0]);
					desired = Double.parseDouble(currentProfileData.getVarValues("desired_proc_units")[0]);
					if (active!=desired)
						map = map | YELLOW;			
				}
				addNumber(sheet, 7, row, lpar[j].getVarValues("run_proc_units"),0, formatFloat(map));
				map = B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW;
				
				addNumber(sheet, 8, row, lpar[j].getVarValues("curr_max_proc_units"),0, formatFloat(map));
				addNumber(sheet, 9, row, lpar[j].getVarValues("run_uncap_weight"),0, formatInt(map));
				addLabel(sheet, 10, row, lpar[j].getVarValues("curr_sharing_mode"),0, formatLabel(map));				
				
				s = lpar[j].getVarValues("curr_shared_proc_pool_name");
				
				if (s!=null) {
					if (s[0].length()>poolNameSize)
						poolNameSize=s[0].length();
					pool = managedSystem[i].getObjects(PROC_POOL);
					for (k=0; k<pool.length; k++)
						if (pool[k].getVarValues("name")[0].equals(s[0]))
							break;
					
					if (k!=pool.length) {
						addLabel(sheet, 11, row, s, 0, formatLabel(map));
						addNumber(sheet, 12, row, pool[k].getVarValues("curr_reserved_pool_proc_units"),0, formatFloat(map));
						addNumber(sheet, 13, row, pool[k].getVarValues("max_pool_proc_units"),0, formatInt(map));
					}
				}
				
				// O113 --> =IF(C113="ded","Dedicated",IF(E113=0,"off",IF(E113>1,ROUND(E113/H113*100,0),"MicroLPAR")))
				f = new Formula(14, row, "IF(C"+(row+1)+"=\"ded\",\"Dedicated\",IF(E"+(row+1)+"=0,\"Off\",IF(E"+(row+1)+">1,ROUND(E"+(row+1)+"/H"+(row+1)+"*100,0),\"MicroLPAR\")))",formatFloat(map));
				sheet.addCell(f); 
				
				addLabel(sheet, 15, row, managedSystem[i].getVarValues("name"),0, formatLabel(map));
				addLabel(sheet, 16, row, managedSystem[i].getVarValues("serial_num"),0, formatLabel(map));
							
				row++;					
			}		
		}

		
		sheet.setColumnView(0, lparNameSize+2);
		sheet.setColumnView(11, poolNameSize+2);
		sheet.setColumnView(3, 10);
		sheet.setColumnView(4, 10);
		sheet.setColumnView(5, 10);
		sheet.setColumnView(6, 10);
		sheet.setColumnView(7, 10);
		sheet.setColumnView(8, 10);
		sheet.setColumnView(9, 10);
		sheet.setColumnView(10, 24);
		sheet.setColumnView(12, 10);
		sheet.setColumnView(13, 10);
		sheet.setColumnView(14, 10);
		sheet.setColumnView(15, 26);
		sheet.setColumnView(16, 26);;
	}
	
	
	
	private void createCPUSheet(WritableSheet sheet)  throws RowsExceededException, WriteException {
		GenericData lpar[], pool[];
		int row;
		int i,j,k;
		String s[];
		Label label;

		int lparNameSize=0;
		int poolNameSize=0;
		
		int ded=0;
		int vp=0;
		int ent100=0;
		
		Formula f;
		
		
		row = 0;
		
		
		
		

		for (i=0; i<managedSystem.length; i++) {
			
			ded=vp=0;
			ent100=0;
			
			/*
			 * Show start of system
			 */
			s = managedSystem[i].getVarValues("name");
			sheet.mergeCells(0, row, 14, row+1);
			addLabel(sheet,0,row,s[0],formatLabel(BOLD|CENTRE|VCENTRE|GREEN));
			row++;
			row++;
			
			
			/*
			 * Setup titles
			 */ 
			sheet.mergeCells(0, row, 0, row+1);
			addLabel(sheet,0,row,"Name",formatLabel(BOLD|B_ALL_MED));
			
			sheet.mergeCells(1, row, 1, row+1);
			addLabel(sheet,1,row,"Status",formatLabel(BOLD|B_ALL_MED));
			
			sheet.mergeCells(2, row, 2, row+1);
			addLabel(sheet,2,row,"Mode",formatLabel(BOLD|B_ALL_MED));
			
			sheet.mergeCells(3, row, 5, row);
			addLabel(sheet,3,row,"Virt/Phys procs",formatLabel(BOLD|CENTRE|B_ALL_MED));
			
			addLabel(sheet,3,row+1,"Min",formatLabel(BOLD|RIGHT|B_ALL_MED));
			addLabel(sheet,4,row+1,"Curr",formatLabel(BOLD|RIGHT|B_ALL_MED));
			addLabel(sheet,5,row+1,"Max",formatLabel(BOLD|RIGHT|B_ALL_MED));
			
			sheet.mergeCells(6, row, 8, row);
			addLabel(sheet,6,row,"Entitlement",formatLabel(BOLD|CENTRE|B_ALL_MED));
			
			addLabel(sheet,6,row+1,"Min",formatLabel(BOLD|RIGHT|B_ALL_MED));
			addLabel(sheet,7,row+1,"Curr",formatLabel(BOLD|RIGHT|B_ALL_MED));
			addLabel(sheet,8,row+1,"Max",formatLabel(BOLD|RIGHT|B_ALL_MED));
			
			sheet.mergeCells(9, row, 9, row+1);
			addLabel(sheet,9,row,"Weight",formatLabel(BOLD|CENTRE|VCENTRE|B_ALL_MED));
			
			sheet.mergeCells(10, row, 10, row+1);
			addLabel(sheet,10,row,"Sharing Mode",formatLabel(BOLD|CENTRE|VCENTRE|B_ALL_MED));
			
			sheet.mergeCells(11, row, 13, row);
			addLabel(sheet,11,row,"Shared Pool",formatLabel(BOLD|CENTRE|B_ALL_MED));
		
			addLabel(sheet,11,row+1,"Name",formatLabel(BOLD|B_ALL_MED));
			addLabel(sheet,12,row+1,"Resv",formatLabel(BOLD|RIGHT|B_ALL_MED));
			addLabel(sheet,13,row+1,"Max",formatLabel(BOLD|RIGHT|B_ALL_MED));
			
			sheet.mergeCells(14, row, 14, row+1);
			addLabel(sheet,14,row,"E:VP %",formatLabel(BOLD|CENTRE|VCENTRE|B_ALL_MED));
			
			row +=2;
			
			
			/*
			 * Write variables
			 */
			lpar = managedSystem[i].getObjects(PROC_LPAR);
			if (lpar==null) {
				row +=2;
				continue;
			}
			
			for (j=0; j<lpar.length; j++) {
				
				s = lpar[j].getVarValues("lpar_name");
				addLabel(sheet, 0, row, s, 0, formatLabel(B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW));
				if (s[0].length()>lparNameSize)
					lparNameSize=s[0].length();
				
				if (lpar[j].getVarValues("run_procs")[0].equals("0"))
					label = new Label(1, row, "Off", formatLabel(B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW));
				else
					label = new Label(1, row, "On", formatLabel(B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW));
				sheet.addCell(label);
				
				addLabel(sheet, 2, row, lpar[j].getVarValues("curr_proc_mode"),0, formatLabel(B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW));
				addNumber(sheet, 3, row, lpar[j].getVarValues("curr_min_procs"),0, formatInt(B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW));
				addNumber(sheet, 4, row, lpar[j].getVarValues("run_procs"),0, formatInt(B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW));
				addNumber(sheet, 5, row, lpar[j].getVarValues("curr_max_procs"),0, formatInt(B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW));
				addNumber(sheet, 6, row, lpar[j].getVarValues("curr_min_proc_units"),0, formatFloat(B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW));
				addNumber(sheet, 7, row, lpar[j].getVarValues("run_proc_units"),0, formatFloat(B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW));
				addNumber(sheet, 8, row, lpar[j].getVarValues("curr_max_proc_units"),0, formatFloat(B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW));
				addNumber(sheet, 9, row, lpar[j].getVarValues("run_uncap_weight"),0, formatInt(B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW));
				addLabel(sheet, 10, row, lpar[j].getVarValues("curr_sharing_mode"),0, formatLabel(B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW));
				
				if (lpar[j].getVarValues("curr_proc_mode")[0].equals("shared")) {
					//vp += Integer.parseInt(lpar[j].getVarValues("run_procs")[0]);
					vp += textToInt(lpar[j].getVarValues("run_procs")[0]);
					ent100 += 100*Float.parseFloat(lpar[j].getVarValues("run_proc_units")[0]);
				} else
					//ded += Integer.parseInt(lpar[j].getVarValues("run_procs")[0]);
					ded += textToInt(lpar[j].getVarValues("run_procs")[0]);
				
				
				s = lpar[j].getVarValues("curr_shared_proc_pool_name");
				//if (s==null) {
				//	row++;
				//	continue;
				//}
				
				if (s!=null) {
					if (s[0].length()>poolNameSize)
						poolNameSize=s[0].length();
					pool = managedSystem[i].getObjects(PROC_POOL);
					for (k=0; k<pool.length; k++)
						if (pool[k].getVarValues("name")[0].equals(s[0]))
							break;
					//if (k==pool.length) {
					//	row++;
					//	continue;	
					//}
					
					if (k!=pool.length) {
						addLabel(sheet, 11, row, s, 0, formatLabel(B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW));
						addNumber(sheet, 12, row, pool[k].getVarValues("curr_reserved_pool_proc_units"),0, formatFloat(B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW));
						addNumber(sheet, 13, row, pool[k].getVarValues("max_pool_proc_units"),0, formatInt(B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW));
					}
				}
				
				// O113 --> =IF(C113="ded","Dedicated",IF(E113=0,"off",IF(E113>1,ROUND(E113/H113*100,0),"MicroLPAR")))
				f = new Formula(14, row, "IF(C"+(row+1)+"=\"ded\",\"Dedicated\",IF(E"+(row+1)+"=0,\"Off\",IF(E"+(row+1)+">1,ROUND(E"+(row+1)+"/H"+(row+1)+"*100,0),\"MicroLPAR\")))",formatFloat(B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW));
				sheet.addCell(f); 
							
				row++;				
			}
			
			
			double sysProcs;
			
			row++;
			
			addLabel(sheet,3,row,"Size",formatLabel(BOLD|CENTRE));
			addLabel(sheet,4,row,"Assigned",formatLabel(BOLD|CENTRE));
			addLabel(sheet,5,row,"Available",formatLabel(BOLD|CENTRE));
			row++;
			
			sheet.mergeCells(0, row, 2, row);
			addLabel(sheet,0,row,"Active Physical Cores",formatLabel(BOLD));
			sysProcs = Double.parseDouble(managedSystem[i].getObjects(PROC)[0].getVarValues("configurable_sys_proc_units")[0]);
			sheet.addCell(new Number(3,row,sysProcs,formatInt(CENTRE|B_TOP_MED|B_LEFT_MED|B_BOTTOM_LOW|B_RIGHT_LOW)));		
			addLabel(sheet,4,row,"",formatLabel(CENTRE|VCENTRE|B_TOP_MED|GRAY_25));
			addLabel(sheet,5,row,"",formatLabel(CENTRE|VCENTRE|B_TOP_MED|B_RIGHT_MED|GRAY_25));
			row++;
			
			sheet.mergeCells(0, row, 2, row);
			addLabel(sheet,0,row,"Dedicated Cores",formatLabel(BOLD));
			addLabel(sheet,3,row,"",formatLabel(CENTRE|VCENTRE|B_LEFT_MED|GRAY_25));			
			sheet.addCell(new Number(4,row,ded,formatInt(CENTRE|VCENTRE|B_ALL_LOW)));
			addLabel(sheet,5,row,"",formatLabel(CENTRE|VCENTRE|B_RIGHT_MED|GRAY_25));			
			row++;
			
			sheet.mergeCells(0, row, 2, row);
			addLabel(sheet,0,row,"Shared Pool",formatLabel(BOLD));
			sheet.addCell(new Number(3,row,sysProcs-ded,formatInt(CENTRE|VCENTRE|B_LEFT_MED|B_TOP_LOW|B_BOTTOM_LOW)));
			sheet.addCell(new Number(4,row,1d*ent100/100,formatFloat(CENTRE|VCENTRE|B_ALL_LOW)));
			sheet.addCell(new Number(5,row,sysProcs-ded-ent100/100d,formatFloat(CENTRE|VCENTRE|B_RIGHT_MED|B_TOP_LOW|B_BOTTOM_LOW)));
			row++;
			
			sheet.mergeCells(0, row, 2, row);
			addLabel(sheet,0,row,"Virtual Processors",formatLabel(BOLD));
			addLabel(sheet,3,row,"",formatLabel(CENTRE|VCENTRE|B_LEFT_MED|B_BOTTOM_MED|GRAY_25));
			sheet.addCell(new Number(4,row,vp,formatInt(CENTRE|VCENTRE|B_BOTTOM_MED|B_LEFT_LOW|B_RIGHT_LOW)));
			addLabel(sheet,5,row,"",formatLabel(CENTRE|VCENTRE|B_RIGHT_MED|B_BOTTOM_MED|GRAY_25));
			row++;

			
			row +=2;
		}

		
		sheet.setColumnView(0, lparNameSize+2);
		sheet.setColumnView(11, poolNameSize+2);
		sheet.setColumnView(3, 10);
		sheet.setColumnView(4, 10);
		sheet.setColumnView(5, 10);
		sheet.setColumnView(6, 10);
		sheet.setColumnView(7, 10);
		sheet.setColumnView(8, 10);
		sheet.setColumnView(9, 10);
		sheet.setColumnView(10, 24);
		sheet.setColumnView(12, 5);
		sheet.setColumnView(13, 7);
		sheet.setColumnView(14, 10);
	}
	
	
	
	private void createSWSheet(WritableSheet sheet)  throws RowsExceededException, WriteException {
		GenericData lpar[], pool[];
		int row;
		int i,j,k;
		String s[];
		Label label;
		GenericData gd[] = null;
		Number number;
		String str;

		
		WritableFont arial10font = new WritableFont(WritableFont.ARIAL, 10);
		WritableFont arial10boldfont = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD, true);
		WritableCellFormat arial10format = new WritableCellFormat (arial10font);
		WritableCellFormat arial10boldformat = new WritableCellFormat (arial10boldfont);
		
		WritableCellFormat centre = new WritableCellFormat (arial10boldfont);
		centre.setAlignment(Alignment.CENTRE);
		centre.setVerticalAlignment(VerticalAlignment.CENTRE);
		
		WritableCellFormat centreGreen = new WritableCellFormat (arial10boldfont);
		centreGreen.setAlignment(Alignment.CENTRE);
		centreGreen.setVerticalAlignment(VerticalAlignment.CENTRE);
		centreGreen.setBackground(Colour.LIGHT_GREEN);
		
		WritableCellFormat title = new WritableCellFormat (arial10boldfont);
		title.setWrap(true);
		
		WritableCellFormat integer = new WritableCellFormat (NumberFormats.INTEGER);
		WritableCellFormat float2d = new WritableCellFormat (NumberFormats.FORMAT3);
		
		Formula f;
		
		
		int lparNameSize=0;
		int poolNameSize=0;
		
		
		
		
		row = 0;
		
		
		
		

		for (i=0; i<managedSystem.length; i++) {
			
			/*
			 * Show start of system
			 */
			s = managedSystem[i].getVarValues("name");
			sheet.mergeCells(0, row, 17, row);
			addLabel(sheet,0,row,s[0],formatLabel(BOLD|CENTRE|VCENTRE|GREEN));
			row++;
			
			/*
			 * Show system cores
			 */
			addLabel(sheet,0,row,"System cores",formatLabel(BOLD|B_ALL_MED));
			
			gd = managedSystem[i].getObjects(PROC);
			if (gd==null || gd[0]==null) {
				addLabel(sheet, 1, row, "N/A", formatLabel(B_RIGHT_LOW|B_BOTTOM_LOW|B_TOP_LOW));
				row +=2;
				continue;
			}
			addNumber(sheet, 1, row, gd[0].getVarValues("configurable_sys_proc_units"),0, formatInt(B_RIGHT_LOW|B_BOTTOM_LOW|B_TOP_LOW));
			row++;
			
			/*
			 * Show processor pool configuration (max 16 pools)
			 */
			int poolsRow = row;		// identifies the row with pool data
			pool = managedSystem[i].getObjects(PROC_POOL);
			sheet.mergeCells(0, row, 0, row+2);
			if (pool==null)
				addLabel(sheet,0,row,"Processor Pools\n(Not supported)",formatLabel(VCENTRE|BOLD|B_ALL_MED|WRAP));
			else
				addLabel(sheet,0,row,"Processor Pools",formatLabel(VCENTRE|BOLD|B_ALL_MED|WRAP));
			/*
			sheet.mergeCells(0, row, 0, row+1);
			label = new Label(0,row,"Processor Pools",centre);
			sheet.addCell(label);
			if (pool==null) {
				// Pools are not supported
				label = new Label(0,row+2,"(not supported)",centre);	
				sheet.addCell(label);
			}
			*/
						
			addLabel(sheet,1,row,"ID",formatLabel(BOLD|B_ALL_MED));
			addLabel(sheet,1,row+1,"Name",formatLabel(BOLD|B_ALL_MED));
			addLabel(sheet,1,row+2,"Size",formatLabel(BOLD|B_ALL_MED));
			/*
			label = new Label(1,row,"ID",arial10boldformat);
			sheet.addCell(label);
			label = new Label(1,row+1,"Name",arial10boldformat);
			sheet.addCell(label);
			label = new Label(1,row+2,"Size",arial10boldformat);
			sheet.addCell(label);
			*/
			
			for (j=1; j<=16; j++) {
				//number = new Number(1+j,row,j,integer);
				//sheet.addCell(number);
				addNumber(sheet, 1+j, row, j, formatInt(B_TOP_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				for (k=0; pool!=null && k<pool.length; k++)
					if (pool[k].getVarValues("shared_proc_pool_id")[0].equals(String.valueOf(j)))
						break;
				if (pool==null || k==pool.length) 
					continue;	
				//label = new Label(1+j,row+1,pool[k].getVarValues("name")[0],arial10format);
				//sheet.addCell(label);
				addLabel(sheet, 1+j, row+1, pool[k].getVarValues("name"), 0, formatLabel(B_RIGHT_LOW|B_BOTTOM_LOW));
				//number = new Number(1+j,row+2,Float.parseFloat(pool[k].getVarValues("max_pool_proc_units")[0]),integer);
				//sheet.addCell(number);	
				addNumber(sheet, 1+j, row+2, pool[k].getVarValues("max_pool_proc_units"), 0, formatInt(B_RIGHT_LOW|B_BOTTOM_LOW));
			}
			row+=4;
			
			
			/*
			 * Setup titles
			 */ 
			sheet.mergeCells(0, row, 0, row+1);
			addLabel(sheet,0,row,"LPAR Name",formatLabel(BOLD|CENTRE|VCENTRE|B_ALL_MED));
			//label = new Label(0,row,"Name",arial10boldformat);
			//sheet.addCell(label);
			
			sheet.mergeCells(1, row, 1, row+1);
			addLabel(sheet,1,row,"Select",formatLabel(BOLD|CENTRE|VCENTRE|B_ALL_MED));
			//label = new Label(1,row,"Select",arial10boldformat);
			//sheet.addCell(label);
			
			sheet.mergeCells(2, row, 2, row+1);
			addLabel(sheet,2,row,"Curr VPs",formatLabel(BOLD|WRAP|CENTRE|VCENTRE|B_ALL_MED));
			//label = new Label(2,row,"Curr",centre);
			//sheet.addCell(label);
			//label = new Label(2,row+1,"VP",centre);
			//sheet.addCell(label);
			
			sheet.mergeCells(3, row, 3, row+1);
			addLabel(sheet,3,row,"Curr Ent",formatLabel(BOLD|WRAP|CENTRE|VCENTRE|B_ALL_MED));
			//label = new Label(3,row,"Curr",centre);
			//sheet.addCell(label);
			//label = new Label(3,row+1,"Ent",centre);
			//sheet.addCell(label);
			
			sheet.mergeCells(4, row, 4, row+1);
			addLabel(sheet,4,row,"cap vs uncap",formatLabel(BOLD|WRAP|CENTRE|VCENTRE|B_ALL_MED));
			//label = new Label(4,row,"cap/uncap",arial10boldformat);
			//sheet.addCell(label);
			
			sheet.mergeCells(5, row, 6, row);
			addLabel(sheet,5,row,"Shared Pool",formatLabel(BOLD|CENTRE|B_ALL_MED));
			//label = new Label(5,row,"Shared Pool",arial10boldformat);
			//sheet.addCell(label);
			
			addLabel(sheet,5,row+1,"ID",formatLabel(BOLD|CENTRE|B_ALL_MED));
			//label = new Label(5,row+1,"ID",arial10boldformat);
			//sheet.addCell(label);
			addLabel(sheet,6,row+1,"Name",formatLabel(BOLD|B_ALL_MED));
			//label = new Label(6,row+1,"Name",arial10boldformat);
			//sheet.addCell(label);
			
			sheet.mergeCells(7, row, 7, row+1);
			addLabel(sheet,7,row,"Max core usage",formatLabel(BOLD|CENTRE|WRAP|B_ALL_MED));
			//label = new Label(7,row,"Max core usage",arial10boldformat);
			//sheet.addCell(label);
			
			sheet.mergeCells(8, row, 25, row);
			addLabel(sheet,8,row,"Max core usage assigned to pools",formatLabel(BOLD|B_ALL_MED));
			//label = new Label(8,row,"Max core usage assigned to pools",arial10boldformat);
			//sheet.addCell(label);
			
			addLabel(sheet,8,row+1,"DedLPAR",formatLabel(BOLD|B_ALL_MED));
			//label = new Label(8,row+1,"DedLPAR",arial10format);
			//sheet.addCell(label);
			addLabel(sheet,9,row+1,"Def Pool",formatLabel(BOLD|B_ALL_MED));
			//label = new Label(9,row+1,"DefaultPool",arial10format);
			//sheet.addCell(label);
			
			for (j=1; j<=16; j++) {
				f = new Formula(9+j, row+1, "IF(ISBLANK("+getColName(1+j)+(poolsRow+2)+"),\"\","+getColName(1+j)+(poolsRow+2)+")",formatLabel(BOLD|B_ALL_MED));
				sheet.addCell(f);
			}
			
			row+=2;
			
			
			/*
			 * Write variables
			 */
			lpar = managedSystem[i].getObjects(PROC_LPAR);
			if (lpar==null) {
				row +=2;
				continue;
			}
			
			boolean dedicated;
			int		lparRow = row;
			
			for (j=0; j<lpar.length; j++) {
				
				s = lpar[j].getVarValues("lpar_name");
				addLabel(sheet, 0, row, s, 0, formatLabel(B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW));
				if (s[0].length()>lparNameSize)
					lparNameSize=s[0].length();
				
				if (lpar[j].getVarValues("curr_proc_mode")[0].equals("ded"))
					dedicated = true;
				else
					dedicated = false;
					
				addNumber(sheet, 2, row, lpar[j].getVarValues("run_procs"),0, formatInt(B_RIGHT_LOW|B_BOTTOM_LOW));
				addNumber(sheet, 3, row, lpar[j].getVarValues("run_proc_units"),0, formatFloat(B_RIGHT_LOW|B_BOTTOM_LOW));
				
				if (!dedicated)
					addLabel(sheet, 4, row, lpar[j].getVarValues("curr_sharing_mode"),0, formatLabel(B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW));
				
				if (dedicated) {
					//number = new Number(5,row,Float.parseFloat("-1"),integer);
					//sheet.addCell(number);
					addNumber(sheet,5,row,-1,formatInt(B_RIGHT_LOW|B_BOTTOM_LOW));
				} else {
					s = lpar[j].getVarValues("curr_shared_proc_pool_id");
					if (s==null) {
						// HMC does not know sub pools (very old one)
						//number = new Number(5,row,Float.parseFloat("0"),integer);
						//sheet.addCell(number);
						addNumber(sheet,5,row,0,formatInt(B_RIGHT_LOW|B_BOTTOM_LOW));
					} else
						addNumber(sheet, 5, row, s,0, formatInt(B_RIGHT_LOW|B_BOTTOM_LOW));
				}

				f = new Formula(6, row, "IF(F"+(row+1)+"=0,\"DefaultPool\",IF(F"+(row+1)+"<0,\"DedLPAR\",HLOOKUP(F"+(row+1)+",$C$"+(poolsRow+1)+":$R$"+(poolsRow+2)+",2)))",formatFloat(B_RIGHT_LOW|B_BOTTOM_LOW));
				//f.setCellFormat(float2d);
				sheet.addCell(f);

				f = new Formula(7, row, "IF(ISBLANK(B"+(row+1)+"),0,IF(OR(F"+(row+1)+"<0,E"+(row+1)+"=\"uncap\"),C"+(row+1)+",D"+(row+1)+"))",formatFloat(B_RIGHT_LOW|B_BOTTOM_LOW));
				//f.setCellFormat(float2d);
				sheet.addCell(f);
				
				f = new Formula(8, row, "IF(F"+(row+1)+"<0,H"+(row+1)+",0)",formatFloat(B_RIGHT_LOW|B_BOTTOM_LOW));
				//f.setCellFormat(float2d);
				sheet.addCell(f);
				
				f = new Formula(9, row, "IF(F"+(row+1)+"=0,H"+(row+1)+",0)",formatFloat(B_RIGHT_LOW|B_BOTTOM_LOW));
				//f.setCellFormat(float2d);
				sheet.addCell(f);
				
				for (k=1; k<=16; k++) {
					f = new Formula(9+k, row, "IF($F"+(row+1)+"="+getColName(1+k)+"$"+(poolsRow+1)+",$H"+(row+1)+",0)",formatFloat(B_RIGHT_LOW|B_BOTTOM_LOW));
					//f.setCellFormat(float2d);
					sheet.addCell(f);
				}	
							
				row++;				
			}
			
			
			f = new Formula(8, row, "SUM(I"+(lparRow+1)+":I"+(row-1+1),formatFloat(B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW));
			//f.setCellFormat(float2d);
			sheet.addCell(f);
			
			f = new Formula(9, row, "MIN(C"+(row+2)+",SUM(J"+(lparRow+1)+":J"+(row-1+1),formatFloat(B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW));
			//f.setCellFormat(float2d);
			sheet.addCell(f);
			
			//=MIN(C27;SUM(K31:K38))
			for (j=10; j<10+16; j++) {
				f = new Formula(j, row, "MIN("+getColName(2+j-10)+(poolsRow+3)+",SUM("+getColName(j)+(lparRow+1)+":"+getColName(j)+(row-1+1),formatFloat(B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW));
				//f.setCellFormat(float2d);
				sheet.addCell(f);
			}
			row++;
			
			label = new Label(0,row,"System pool size",formatLabel(BOLD|B_ALL_MED));
			sheet.addCell(label);
			f = new Formula(2, row, "B"+(poolsRow-1+1)+"-I"+(row));
			f.setCellFormat(float2d);
			sheet.addCell(f);
			row++;
			
			label = new Label(0,row,"Max dedicated cores",formatLabel(BOLD|B_ALL_MED));
			sheet.addCell(label);
			f = new Formula(2, row, "I"+(row-2+1));
			f.setCellFormat(float2d);
			sheet.addCell(f);
			row++;
			
			label = new Label(0,row,"Max pool cores",formatLabel(BOLD|B_ALL_MED));
			sheet.addCell(label);
			f = new Formula(2, row, "MIN(C"+(row-2+1)+",SUM(J"+(row-3+1)+":Z"+(row-3+1));
			f.setCellFormat(float2d);
			sheet.addCell(f);
			row++;

			label = new Label(0,row,"Max used cores",formatLabel(BOLD|B_ALL_MED));
			sheet.addCell(label);
			f = new Formula(2, row, "C"+(row-2+1)+"+C"+(row-1+1));
			f.setCellFormat(float2d);
			sheet.addCell(f);
			label = new Label(3,row,"<<<=== Useful when checking number of required SW licenses",arial10format);
			sheet.addCell(label);
			row++;
						
			
			row +=2;
		}

		if (lparNameSize<=18)
			lparNameSize=20;
		sheet.setColumnView(0, lparNameSize+2);

	}
	
	
	
	private int addLabel(WritableSheet sheet, int col, int row, String s, WritableCellFormat format) throws RowsExceededException, WriteException {
		if (s==null){
			sheet.addCell(new Label(col,row,null,format));
			return 0;
		}
		
		sheet.addCell(new Label(col,row,s,format));
		return s.length();
	}
	

	
	private int  addLabel(WritableSheet sheet, int col, int row, String s[], int index, WritableCellFormat format) throws RowsExceededException, WriteException {
		if ( s==null || index>=s.length) {
			sheet.addCell(new Label(col,row,null,format));
			return 0;
		}
		
		sheet.addCell(new Label(col,row,s[index],format));
		if (s[index]==null)
			return 0;
		return s[index].length();
	}
	
	
	private void addMultipleLabels(WritableSheet sheet, int col, int row, String s[], WritableCellFormat format) throws RowsExceededException, WriteException {
		if (sheet==null || s==null)
			return;
		
		if (s.length==1)
			sheet.addCell(new Label(col,row,s[0],format));
		else {		
			String result = s[0];
			
			for (int i=1; i<s.length; i++)
				result = result + ", " + s[i];	
			sheet.addCell(new Label(col,row,result,format));
		}
	}
	
	private void addMultipleLabelsWrap(WritableSheet sheet, int col, int row, String s[], WritableCellFormat format) throws RowsExceededException, WriteException {
		if (sheet==null || s==null)
			return;
		
		if (s.length==0)
			return;
		
		if (s.length==1)
			sheet.addCell(new Label(col,row,s[0],format));
		else {		
			String result = s[0];
			
			for (int i=1; i<s.length; i++)
				//result = result + ", " + s[i];	
				result = result + "\n" + s[i];
			sheet.addCell(new Label(col,row,result,format));
		}
	}
	
	
	private void addNumber(WritableSheet sheet, int col, int row, String s[], int index, WritableCellFormat format) throws RowsExceededException, WriteException {		
		if (sheet==null || s==null)
			return;
		
		if (index>=s.length)
			return;
		
		double d; 
		
		try {
			d = Double.parseDouble(s[index]);
			
			// Only keep 2 digits: avoid rounding errors
			d = 1d*(int)(d*100)/100;
			
			sheet.addCell(new Number(col,row,d,format));
		} catch (NumberFormatException nfe) {}	
	}
	
	private void addNumber(WritableSheet sheet, int col, int row, double d, WritableCellFormat format) throws RowsExceededException, WriteException {
		sheet.addCell(new Number(col,row,d,format));
	}
	
	private void addNumberDiv1024(WritableSheet sheet, int col, int row, String s[], int index, WritableCellFormat format) throws RowsExceededException, WriteException {		
		if (sheet==null || s==null)
			return;
		
		if (index>=s.length)
			return;
		
		double d; 
		
		try {
			d = Double.parseDouble(s[index]);
			d = d / 1024;
			
			// Only keep 2 digits: avoid rounding errors
			d = 1d*(int)(d*100)/100;
			
			sheet.addCell(new Number(col,row,d,format));
		} catch (NumberFormatException nfe) {}	
	}

	
	
	private void createSlotsSheet(WritableSheet sheet) throws RowsExceededException, WriteException {
		String systemName;
		GenericData object[];
		int row;
		String names[];
		int i,j,k,x;
		String s[];
		Label label;
		String v;
		
		WritableFont arial10font = new WritableFont(WritableFont.ARIAL, 10);
		WritableFont arial10boldfont = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD, true);
		WritableCellFormat arial10format = new WritableCellFormat (arial10font);
		WritableCellFormat arial10boldformat = new WritableCellFormat (arial10boldfont);
		
		WritableCellFormat centre = new WritableCellFormat (arial10boldfont);
		centre.setAlignment(Alignment.CENTRE);
		centre.setVerticalAlignment(VerticalAlignment.CENTRE);
		centre.setBackground(Colour.LIGHT_GREEN);
		
		row = 0;
		
		// Merge variable names
		names=null;
		for (i=0; i<managedSystem.length; i++) {		
			object = managedSystem[i].getObjects(SLOT);
			if (object==null)
				continue;			
			for (j=0; j<object.length; j++)  {
				s = object[j].getVarNames();
				if (names==null)
					names=s;
				else
					names = mergeList(names, s);
			}
		}
		
		
		
		int columnSize[] = new int[names.length];
		for (i=0; i<columnSize.length; i++)
			columnSize[i]=names[i].length();
		
		
		row=0;
		for (i=0; i<managedSystem.length; i++) {
			systemName = managedSystem[i].getVarValues("name")[0];
			
			sheet.mergeCells(0, row, names.length-1, row);
			label = new Label(0,row,systemName+" Slots",centre);
			sheet.addCell(label);
			row++;
			
			object = managedSystem[i].getObjects(SLOT);
			if (object==null) {
				row+=2;
				continue;
			}
			
			// Print variable names				
			for (j=0; j<names.length; j++) {
				label = new Label(j,row,names[j],arial10boldformat);
				sheet.addCell(label);					
			}
			row++;
			
			for (j=0; j<object.length; j++) {
				for (k=0; k<names.length; k++) {
					s = object[j].getVarValues(names[k]);
					if (s==null)
						continue;
					
					v=null;
					for (x=0; x<s.length; x++) {	
						if (v==null)
							v=s[x];
						else
							v=v+","+s[x];
					}
				
					label = new Label(k,row,v,arial10format);
					sheet.addCell(label);
					
					if (v.length()>columnSize[k])
						columnSize[k]=v.length();
				}
				row++;
			}			
			row+=2;						
		}	
		
		for (i=0; i<columnSize.length; i++)
			sheet.setColumnView(i, columnSize[i]+2);
	}	
	
	


	
	private void createHeader(WritableSheet sheet) throws RowsExceededException, WriteException {
		int row;
			
		row = 0;
		
		addLabel(sheet,0,row,"Manager Name",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,1,row,scannerParams.getVarValues("HMC")[0],formatLabel(B_ALL_LOW));
		row++;
		
		addLabel(sheet,0,row,"Manager Type",formatLabel(BOLD|B_ALL_MED));
		switch (managerType) {
			case M_FSM: 	addLabel(sheet,1,row,"FSM",formatLabel(B_ALL_LOW)); break;
			case M_HMC: 	addLabel(sheet,1,row,"HMC",formatLabel(B_ALL_LOW)); break;
			case M_SDMC: 	addLabel(sheet,1,row,"SDMC",formatLabel(B_ALL_LOW)); break;
			case M_IVM: 	addLabel(sheet,1,row,"IVM",formatLabel(B_ALL_LOW)); break;
			default:		break;
		}
		row++;
		
		addLabel(sheet,0,row,"User",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,1,row,scannerParams.getVarValues("user")[0],formatLabel(B_ALL_LOW));
		row++;
		
		addLabel(sheet,0,row,"Generation Date",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,1,row,scannerParams.getVarValues("date")[0],formatLabel(B_ALL_LOW));
		row++;
		
		addLabel(sheet,0,row,"Generation Time",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,1,row,scannerParams.getVarValues("time")[0],formatLabel(B_ALL_LOW));
		row++;
		
		addLabel(sheet,0,row,"Manager Date",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,1,row,scannerParams.getVarValues("HMCdate")[0],formatLabel(B_ALL_LOW));
		row++;
		
		addLabel(sheet,0,row,"Manager Time",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,1,row,scannerParams.getVarValues("HMCtime")[0],formatLabel(B_ALL_LOW));
		row++;
		
		addLabel(sheet,0,row,"HMC Scanner Version",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,1,row,version,formatLabel(B_ALL_LOW));
		row++;
		
		
		row++;
		addLabel(sheet,0,row,"Latest HMC Scanner is available at http://tinyurl.com/HMCscanner",formatLabel(NONE));
		row++;
		addLabel(sheet,0,row,"For issues send a mail to vagnini@it.ibm.com",formatLabel(NONE));
		row++;
		
		
		File f = new File(System.getProperty("user.dir")+File.separatorChar+hmcScannerPic);
		if (f.exists()) {
			WritableImage imgobj=new WritableImage(3, 2, 5, 15, f);
			sheet.addImage(imgobj);
		}
		
		
		
		sheet.setColumnView(0, 21);
		sheet.setColumnView(1, 36);
		
	}
	
	
	
	private void createHeaderExcel(WritableSheet sheet) {
		DataSheet ds = createHeader();
		if (ds!=null)
			ds.createExcelSheet(sheet);
	}
	
	private void createHeaderHTML(String fileName) {
		DataSheet ds = createHeader();
		if (ds!=null) {
			ds.createHTMLSheet(fileName);
			addButton("Header",new File(fileName).getName());
		}
	}
	
	private void createHeaderCSV(String fileName) {
		DataSheet ds = createHeader();
		if (ds!=null) {
			ds.setSeparator(csvSeparator);
			ds.createCSVSheet(fileName);
		}
	}
	
	private DataSheet createHeader() {
		DataSheet ds = new DataSheet();
		int row;
			
		row = 0;
		
		ds.addLabel(0,row,"Manager Name",BOLD|B_ALL_MED);
		ds.addLabel(1,row,scannerParams.getVarValues("HMC")[0],B_ALL_LOW);
		row++;
		
		ds.addLabel(0,row,"Manager Type",BOLD|B_ALL_MED);
		switch (managerType) {
			case M_FSM: 	ds.addLabel(1,row,"FSM",B_ALL_LOW); break;
			case M_HMC: 	ds.addLabel(1,row,"HMC",B_ALL_LOW); break;
			case M_SDMC: 	ds.addLabel(1,row,"SDMC",B_ALL_LOW); break;
			case M_IVM: 	ds.addLabel(1,row,"IVM",B_ALL_LOW); break;
			default:		break;
		}
		row++;
		
		ds.addLabel(0,row,"User",BOLD|B_ALL_MED);
		ds.addLabel(1,row,scannerParams.getVarValues("user")[0],B_ALL_LOW);
		row++;
		
		ds.addLabel(0,row,"Generation Date",BOLD|B_ALL_MED);
		ds.addLabel(1,row,scannerParams.getVarValues("date")[0],B_ALL_LOW);
		row++;
		
		ds.addLabel(0,row,"Generation Time",BOLD|B_ALL_MED);
		ds.addLabel(1,row,scannerParams.getVarValues("time")[0],B_ALL_LOW);
		row++;
		
		ds.addLabel(0,row,"Manager Date",BOLD|B_ALL_MED);
		ds.addLabel(1,row,scannerParams.getVarValues("HMCdate")[0],B_ALL_LOW);
		row++;
		
		ds.addLabel(0,row,"Manager Time",BOLD|B_ALL_MED);
		ds.addLabel(1,row,scannerParams.getVarValues("HMCtime")[0],B_ALL_LOW);
		row++;
		
		ds.addLabel(0,row,"HMC Scanner Version",BOLD|B_ALL_MED);
		ds.addLabel(1,row,version,B_ALL_LOW);
		row++;
		
		
		row++;
		ds.addLabel(0,row,"Latest HMC Scanner is available at http://tinyurl.com/HMCscanner",NONE);
		row++;
		ds.addLabel(0,row,"For issues send a mail to vagnini@it.ibm.com",NONE);
		row++;
		
		
		ds.addPicture(3, row, 5, 15, System.getProperty("user.dir")+File.separatorChar+hmcScannerPic);
		
		ds.setColSize(0, 21);
		ds.setColSize(1, 36);	
		
		return ds;
	}
	
	private void createSystemsSheet2(WritableSheet sheet) throws RowsExceededException, WriteException {
		int i;
		GenericData gd[];
		int row;
		int nameSize=0;
		String s;
			
		row = 0;
		
		/*
		 * Setup titles
		 */ 
		sheet.mergeCells(0, row, 0, row+1);
		addLabel(sheet,0,row,"Name",formatLabel(BOLD|VCENTRE|B_ALL_MED));
		
		sheet.mergeCells(1, row, 1, row+1);
		addLabel(sheet,1,row,"Status",formatLabel(BOLD|VCENTRE|B_ALL_MED));
		
		sheet.mergeCells(2, row, 3, row);
		addLabel(sheet,2,row,"Identification",formatLabel(BOLD|CENTRE|B_ALL_MED));
		
		addLabel(sheet,2,row+1,"Type-Model",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,3,row+1,"Serial",formatLabel(BOLD|B_ALL_MED));
		
		sheet.mergeCells(4, row, 8, row);
		addLabel(sheet,4,row,"Cores",formatLabel(BOLD|CENTRE|B_ALL_MED));
		
		addLabel(sheet,4,row+1,"Installed",formatLabel(BOLD|B_ALL_MED|RIGHT));
		addLabel(sheet,5,row+1,"Active",formatLabel(BOLD|B_ALL_MED|RIGHT));
		addLabel(sheet,6,row+1,"Deconfig",formatLabel(BOLD|B_ALL_MED|RIGHT));
		addLabel(sheet,7,row+1,"Curr Avail",formatLabel(BOLD|B_ALL_MED|RIGHT));
		addLabel(sheet,8,row+1,"Pend Avail",formatLabel(BOLD|B_ALL_MED|RIGHT));
		
		sheet.mergeCells(9, row, 14, row);
		addLabel(sheet,9,row,"Memory (MB)",formatLabel(BOLD|CENTRE|B_ALL_MED));
		
		addLabel(sheet,9,row+1,"Installed",formatLabel(BOLD|B_ALL_MED|RIGHT));
		addLabel(sheet,10,row+1,"Active",formatLabel(BOLD|B_ALL_MED|RIGHT));
		addLabel(sheet,11,row+1,"Deconfig",formatLabel(BOLD|B_ALL_MED|RIGHT));
		addLabel(sheet,12,row+1,"Firmware",formatLabel(BOLD|B_ALL_MED|RIGHT));
		addLabel(sheet,13,row+1,"Curr Avail",formatLabel(BOLD|B_ALL_MED|RIGHT));
		addLabel(sheet,14,row+1,"Pend Avail",formatLabel(BOLD|B_ALL_MED|RIGHT));
		
		sheet.mergeCells(15, row, 15, row+1);
		addLabel(sheet,15,row,"Data Sec",formatLabel(BOLD|VCENTRE|B_ALL_MED|RIGHT));
		
		row +=2;
		
		
		/*
		 * Show values, each system on one line
		 */
		for (i=0; i<managedSystem.length; i++) {
			s = managedSystem[i].getVarValues("name")[0];
			if (s.length() > nameSize)
				nameSize = s.length();
			
			addLabel(sheet,0,row,s,formatLabel(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
			if (managerType == M_SDMC)
				addLabel(sheet,1,row,managedSystem[i].getVarValues("primary_state")[0],formatLabel(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
			else
				addLabel(sheet,1,row,managedSystem[i].getVarValues("state")[0],formatLabel(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
			addLabel(sheet,2,row,managedSystem[i].getVarValues("type_model")[0],formatLabel(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
			addLabel(sheet,3,row,managedSystem[i].getVarValues("serial_num")[0],formatLabel(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
			
			gd = managedSystem[i].getObjects(PROC);
			if (gd!=null && gd[0]!=null) {		
				addNumber(sheet,4,row,gd[0].getVarValues("installed_sys_proc_units"),0,formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				addNumber(sheet,5,row,gd[0].getVarValues("configurable_sys_proc_units"),0,formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				addNumber(sheet,6,row,gd[0].getVarValues("deconfig_sys_proc_units"),0,formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				addNumber(sheet, 7, row, gd[0].getVarValues("curr_avail_sys_proc_units"),0, formatFloat(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				addNumber(sheet, 8, row, gd[0].getVarValues("pend_avail_sys_proc_units"),0, formatFloat(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
			}
			
			gd = managedSystem[i].getObjects(MEM);
			if (gd!=null && gd[0]!=null) {	
				addNumber(sheet, 9, row, gd[0].getVarValues("installed_sys_mem"),0, formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				addNumber(sheet, 10, row, gd[0].getVarValues("configurable_sys_mem"),0, formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				addNumber(sheet, 11, row, gd[0].getVarValues("deconfig_sys_mem"),0, formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				addNumber(sheet, 12, row, gd[0].getVarValues("sys_firmware_mem"),0, formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				addNumber(sheet, 13, row, gd[0].getVarValues("curr_avail_sys_mem"),0, formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				addNumber(sheet, 14, row, gd[0].getVarValues("pend_avail_sys_mem"),0, formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
			}
			
			addNumber(sheet, 15, row, managedSystem[i].getVarValues("sample_rate"),0, formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
			
			row++;
		}
		
		// Set column size
		for (i=0; i<=15; i++)
			sheet.setColumnView(i, 14);	
		sheet.setColumnView(0, nameSize+2);	
		sheet.setColumnView(1, 12);	
		sheet.setColumnView(2, 12);	
		sheet.setColumnView(3, 12);			
	}
	
	
	private void createSysRAMUsageSheetExcel(WritableSheet sheet) {
		DataSheet ds = createSysRAMUsageSheet();
		if (ds!=null)
			ds.createExcelSheet(sheet);
	}
	
	private void createSysRAMUsageSheetHTML(String fileName) {
		DataSheet ds = createSysRAMUsageSheet();
		if (ds!=null) {
			ds.createHTMLSheet(fileName);
			addButton("System RAM Usage",new File(fileName).getName());
		}
	}
	
	private void createSysRAMUsageSheetCSV(String fileName) {
		DataSheet ds = createSysRAMUsageSheet();
		if (ds!=null) {
			ds.setSeparator(csvSeparator);
			ds.createCSVSheet(fileName);
		}
	}
	
	private DataSheet createSysRAMUsageSheet() {
		DataSheet ds = new DataSheet();
		int i,j;
		String s;
		float f,f2;
		int		size=0;
		int		row;
		
		String timeLabels[] = msMemConfig[0].getMonthlyLabels();

		
		row = 0;
		
		ds.mergeCells(0, row, 13, row);
		ds.addLabel(0,row,"Average Free Memory GB",BOLD|CENTRE|VCENTRE|GREEN);
		ds.mergeCells(15, row, 28, row);
		ds.addLabel(15,row,"Average Percentage Free Memory",BOLD|CENTRE|VCENTRE|GREEN);
		row++;
		
		for (i=0; i<timeLabels.length; i++) {
			s = timeLabels[i];
			ds.addLabel(1+i,row,s,BOLD|B_ALL_MED);
			ds.addLabel(16+i,row,s,BOLD|B_ALL_MED);
		}
		row++;
		
		
		for (i=0; i<managedSystem.length; i++) {
			s = managedSystem[i].getVarValues("name")[0];
			if (s.length()>size)
				size=s.length();
			
			ds.addLabel(0,row+i,s,BOLD|VCENTRE|CENTRE|B_ALL_MED);
			ds.addLabel(15,row+i,s,BOLD|VCENTRE|CENTRE|B_ALL_MED);
			
			for (j=0; j<timeLabels.length; j++) {
				f = msMemAvail[i].getMonthData(j);
				f2 = msMemConfig [i].getMonthData(j);
				if (f>=0)
					ds.addFloat(1+j,row+i, f, B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW);
				if (f>=0 && f2>0)
					ds.addFloat(16+j,row+i, f/f2*100, B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW);
			}
		}
		row+=i;
		row++;
		
		
		ds.mergeCells(0, row, 13, row);
		ds.addLabel(0,row,"Average Configured Memory GB",BOLD|CENTRE|VCENTRE|GREEN);
		row++;
		
		for (i=0; i<timeLabels.length; i++) {
			s = timeLabels[i];
			ds.addLabel(1+i,row,s,BOLD|B_ALL_MED);
		}
		row++;
		
		
		for (i=0; i<managedSystem.length; i++) {
			s = managedSystem[i].getVarValues("name")[0];
			if (s.length()>size)
				size=s.length();
			
			ds.addLabel(0,row+i,s,BOLD|VCENTRE|CENTRE|B_ALL_MED);
			
			for (j=0; j<timeLabels.length; j++) {
				f = msMemConfig [i].getMonthData(j);
				if (f>=0)
					ds.addFloat(1+j,row+i, f, B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW);
			}
		}
		row+=i;
		row++;
		
		
		ds.setColSize(0, size+6);
		ds.setColSize(15, size+6);

		return ds;
	}
	
	
	
	private void createSysRAMUsageSheet(WritableSheet sheet) throws RowsExceededException, WriteException {	
		int i,j;
		String s;
		float f,f2;
		int		size=0;
		int		row;
		
		String timeLabels[] = msMemConfig[0].getMonthlyLabels();
		
		row = 0;
		
		sheet.mergeCells(0, row, 13, row);
		addLabel(sheet,0,row,"Average Free Memory GB",formatLabel(BOLD|CENTRE|VCENTRE|GREEN));
		sheet.mergeCells(15, row, 28, row);
		addLabel(sheet,15,row,"Average Percentage Free Memory",formatLabel(BOLD|CENTRE|VCENTRE|GREEN));
		row++;
		
		for (i=0; i<timeLabels.length; i++) {
			s = timeLabels[i];
			addLabel(sheet,1+i,row,s,formatLabel(BOLD|B_ALL_MED));
			addLabel(sheet,16+i,row,s,formatLabel(BOLD|B_ALL_MED));
		}
		row++;
		
		
		for (i=0; i<managedSystem.length; i++) {
			s = managedSystem[i].getVarValues("name")[0];
			if (s.length()>size)
				size=s.length();
			
			addLabel(sheet,0,row+i,s,formatLabel(BOLD|VCENTRE|CENTRE|B_ALL_MED));
			addLabel(sheet,15,row+i,s,formatLabel(BOLD|VCENTRE|CENTRE|B_ALL_MED));
			
			for (j=0; j<timeLabels.length; j++) {
				f = msMemAvail[i].getMonthData(j);
				f2 = msMemConfig [i].getMonthData(j);
				if (f>=0)
					addNumber(sheet,1+j,row+i, f, formatFloat(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				if (f>=0 && f2>0)
					addNumber(sheet,16+j,row+i, f/f2*100, formatFloat(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
			}
		}
		row+=i;
		row++;
		
		
		sheet.mergeCells(0, row, 13, row);
		addLabel(sheet,0,row,"Average Configured Memory GB",formatLabel(BOLD|CENTRE|VCENTRE|GREEN));
		row++;
		
		for (i=0; i<timeLabels.length; i++) {
			s = timeLabels[i];
			addLabel(sheet,1+i,row,s,formatLabel(BOLD|B_ALL_MED));
		}
		row++;
		
		
		for (i=0; i<managedSystem.length; i++) {
			s = managedSystem[i].getVarValues("name")[0];
			if (s.length()>size)
				size=s.length();
			
			addLabel(sheet,0,row+i,s,formatLabel(BOLD|VCENTRE|CENTRE|B_ALL_MED));
			
			for (j=0; j<timeLabels.length; j++) {
				f = msMemConfig [i].getMonthData(j);
				if (f>=0)
					addNumber(sheet,1+j,row+i, f, formatFloat(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
			}
		}
		row+=i;
		row++;
		
		
		sheet.setColumnView(0, size+6);
		sheet.setColumnView(15, size+6);
	}
	
	
	
	private void createLparDailyUsageSheetExcel(WritableSheet sheet, int from, int to) {
		DataSheet ds = createLparDailyUsageSheet(from, to);
		if (ds!=null)
			ds.createExcelSheet(sheet);
	}
	
	private void createLparDailyUsageSheetHTML(String fileName, int from, int to, int count) {
		DataSheet ds = createLparDailyUsageSheet(from, to);
		if (ds!=null) {
			ds.createHTMLSheet(fileName);
			addButton("LPAR Daily N"+count,new File(fileName).getName());
		}
	}
	
	private void createLparDailyUsageSheetCSV(String fileName, int from, int to, int count) {
		DataSheet ds = createLparDailyUsageSheet(from, to);
		if (ds!=null) {
			ds.setSeparator(csvSeparator);
			ds.createCSVSheet(fileName);
		}
	}
	
	private DataSheet createLparDailyUsageSheet(int from, int to) {
		DataSheet ds = new DataSheet();
		int i,j;
		String s;
		float f;
		int		size=0;
		int		row;
		
		boolean cap;
		String	ms, pool;
		float	vp, ent;
		float	free;
		int		poolID;
		float max;
		int red[], yellow[], num[], uncap[];
		int map=0;
		String col;
		
		GregorianCalendar gc=null;
		
		String timeLabels[] = lparPC[from].getDailyLabels();
		
		if (to>lparNames.length-1)
			to=lparNames.length-1;
		
		
		row=0;	
		
		ds.addLabel(0,row,"Red",BOLD|CENTRE|VCENTRE|RED);
		ds.mergeCells(1, row, 10, row);
		ds.addLabel(1,row,"Usage >= "+R_LEVEL+"% of max allowed size",BOLD);
		row++;
		
		ds.addLabel(0,row,"Yellow",BOLD|CENTRE|VCENTRE|YELLOW);
		ds.mergeCells(1, row, 10, row);
		ds.addLabel(1,row,"Usage >= "+Y_LEVEL+"% of max allowed size",BOLD);
		row++;
		
		ds.addLabel(0,row,"Bold",BOLD|CENTRE|VCENTRE);
		ds.mergeCells(1, row, 10, row);
		ds.addLabel(1,row,"Usage > entitled capacity",BOLD);
		row++;
		
		ds.addLabel(0,row,"LPAR color",BOLD|CENTRE|VCENTRE);
		ds.mergeCells(1, row, 10, row);
		ds.addLabel(1,row,"Colored cells >= "+COLOR_LEVEL*100+"% of total (yellows include reds)",BOLD);
		row++;
		row++;
		
		ds.mergeCells(0, row, to-from+2, row);
		ds.addLabel(0,row,"Average LPAR Usage",BOLD|LEFT|VCENTRE|GREEN);
		row++;
		row++;
		
		ds.addLabel(0,row,"Days",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
		ds.addLabel(0,row,"Reds %",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
		ds.addLabel(0,row,"Yellows %",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
		ds.addLabel(0,row,"Over Cap %",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
		ds.addLabel(0,row,"Max",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
		ds.addLabel(0,row,"Avg",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
		ds.addLabel(0,row,"90% <=",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
		ds.addLabel(0,row,"95% <=",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
		ds.addLabel(0,row,"Last Ent",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
		row++;
		ds.addLabel(0,row,"Date",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
		
		for (i=0; i<timeLabels.length; i++) {
			s = timeLabels[i];
			ds.addLabel(0,row+i,s,BOLD|B_ALL_MED);
		}	
			
		
		red = new int[to-from+1];
		yellow = new int[red.length];
		num = new int[red.length];
		uncap = new int[red.length];
		
		for (i=from; i<=to; i++) {			
			for (j=0; j<365; j++) {
				f = lparPC[i].getDayData(j);
				
				ent = lparEnt[i].getDayData(j);
				cap = lparStatus[i].getDayCap(j);
				ms = lparStatus[i].getDayMS(j);
				pool = lparStatus[i].getDayPool(j);
				vp = lparVP[i].getDayData(j);
				
				if (f<0 /*|| ent<=0 */)
					continue;
				
				if (pool.equals("DefaultPool")) {
					free = msCoreConfig[getMSid(ms)].getDayData(j) - msCoreUsed[getMSid(ms)].getDayData(j);
				} else {
					poolID = getProcPoolId(ms,pool);
					free = procPoolConfig[poolID].getDayData(j)-procPoolUsed[poolID].getDayData(j);
				}
				
				if (cap) {
					max = ent;					
				} else {
					if (f+free>=vp)
						max = vp;
					else
						max = f+free;
					if (max<ent)
						max=ent;
				}
				map = B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW;
				if (f/max*100>=R_LEVEL) {
					map = map|RED;
					red[i-from]++;
					yellow[i-from]++;
				} else if (f/max*100>=Y_LEVEL) {
					map = map|YELLOW;
					yellow[i-from]++;
				}
				
				num[i-from]++;
				
				if (f>ent) {
					map = map | BOLD;
					uncap[i-from]++;
				}
					
				ds.addFloat(1+i-from,row+j, f, map);
				
				/*
				if (f>=0)
					addNumber(sheet,1+i-from,row+j, f, formatFloat(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				*/
			}
		}
		
		// Add lpar and counters
		for (i=from; i<=to; i++) {
			s = lparNames[i];
			if (s.length()>size)
				size=s.length();
			
			map = BOLD|DIAG45|B_ALL_MED;
			if (1d*red[i-from]/num[i-from]>COLOR_LEVEL)
				map = map | RED;
			else if (1d*yellow[i-from]/num[i-from]>COLOR_LEVEL)
				map = map | YELLOW;
			
			ds.addLabel(1+i-from,row-12,s,map);
			
			map = B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW;
			ds.addInteger(1+i-from,row-11, num[i-from], map);
			
			if (num[i-from]==0) {
				map = B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW;
				ds.addFloat(1+i-from,row-10, 0, map);
				ds.addFloat(1+i-from,row-9, 0, map);
				ds.addFloat(1+i-from,row-8, 0, map);
				ds.addFloat(1+i-from,row-7, 0, map);
				ds.addFloat(1+i-from,row-6, 0, map);
				ds.addFloat(1+i-from,row-5, 0, map);
				ds.addFloat(1+i-from,row-4, 0, map);
				ds.addFloat(1+i-from,row-3, 0, map);
			} else {		
				if (1d*red[i-from]/num[i-from]>COLOR_LEVEL)
					map = map | RED;
				ds.addFloat(1+i-from,row-10, 100d*red[i-from]/num[i-from], map);
				
				map = B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW;
				if (1d*yellow[i-from]/num[i-from]>COLOR_LEVEL)
					map = map | YELLOW;
				ds.addFloat(1+i-from,row-9, 100d*yellow[i-from]/num[i-from], map);
				
				map = B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW;
				ds.addFloat(1+i-from,row-8, 100d*uncap[i-from]/num[i-from], map);	
				
				col = jxl.CellReferenceHelper.getColumnReference(1+i-from);
				
				ds.addFloat(1+i-from, row-7, lparPC[i].getDailyMax(), B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
				//ds.addFormula(1+i-from, row-7, "MAX("+col+(row+1)+":"+col+(row+365)+")", B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
								
				ds.addFloat(1+i-from, row-6, lparPC[i].getDailyAvg(), B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
				//ds.addFormula(1+i-from, row-6, "AVERAGE("+col+(row+1)+":"+col+(row+365)+")", B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
				
				ds.addFloat(1+i-from, row-5, lparPC[i].getDaily90p(), B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
				//ds.addFormula(1+i-from, row-5, "STDEVP("+col+(row+1)+":"+col+(row+365)+")", B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
					
				ds.addFloat(1+i-from, row-4, lparPC[i].getDaily95p(), B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
				//ds.addFormula(1+i-from, row-4, "1282/1000*"+col+(row-5+1)+"+"+col+(row-6+1)+")", B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
								
				for (j=364; j>=0; j--)
					if ( lparEnt[i].getDayData(j)>0) break;
				if (j>=0)  // id dedicated LPAR there is no ent
					ds.addFloat(1+i-from,row-3, lparEnt[i].getDayData(j), B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);		
			}
		}
		
		row++;
		
		row+=365;
		row++;	

		ds.setColSize(0, 12);
		for (i=1; i<=(to-from)+1; i++)
			ds.setColSize(i, 7);

		return ds;
	}
	
	
	
	
	
	private void createLparHourlyUsageSheetExcel(WritableSheet sheet, int from, int to) {
		DataSheet ds = createLparHourlyUsageSheet(from, to);
		if (ds!=null)
			ds.createExcelSheet(sheet);
	}
	
	private void createLparHourlyUsageSheetHTML(String fileName, int from, int to, int count) {
		DataSheet ds = createLparHourlyUsageSheet(from, to);
		if (ds!=null) {
			ds.createHTMLSheet(fileName);
			addButton("LPAR Hourly N"+count,new File(fileName).getName());
		}
	}
	
	private void createLparHourlyUsageSheetCSV(String fileName, int from, int to, int count) {
		DataSheet ds = createLparHourlyUsageSheet(from, to);
		if (ds!=null) {
			ds.setSeparator(csvSeparator);
			ds.createCSVSheet(fileName);
		}
	}
	
	private DataSheet createLparHourlyUsageSheet(int from, int to) {
		DataSheet ds = new DataSheet();
		int i,j;
		String s;
		float f;
		int		size=0;
		int		row;
		
		boolean cap;
		String	ms, pool;
		float	vp, ent;
		float	free;
		int		poolID;
		float max;
		int red[], yellow[], num[], uncap[];
		int map=0;
		String col;
		
		GregorianCalendar gc=null;
		
		String timeLabels[] = lparPC[from].getHourlyLabels();
		
		if (to>lparNames.length-1)
			to=lparNames.length-1;
		
		
		row=0;		
		
		ds.addLabel(0,row,"Red",BOLD|CENTRE|VCENTRE|RED);
		ds.mergeCells(1, row, 10, row);
		ds.addLabel(1,row,"Usage >= "+R_LEVEL+"% of max allowed size",BOLD);
		row++;
		
		ds.addLabel(0,row,"Yellow",BOLD|CENTRE|VCENTRE|YELLOW);
		ds.mergeCells(1, row, 10, row);
		ds.addLabel(1,row,"Usage >= "+Y_LEVEL+"% of max allowed size",BOLD);
		row++;
		
		ds.addLabel(0,row,"Bold",BOLD|CENTRE|VCENTRE);
		ds.mergeCells(1, row, 10, row);
		ds.addLabel(1,row,"Usage > entitled capacity",BOLD);
		row++;
		
		ds.addLabel(0,row,"LPAR color",BOLD|CENTRE|VCENTRE);
		ds.mergeCells(1, row, 10, row);
		ds.addLabel(1,row,"Colored cells >= "+COLOR_LEVEL*100+"% of total (yellows include reds)",BOLD);
		row++;
		row++;
		
		ds.mergeCells(0, row, to-from+2, row);
		ds.addLabel(0,row,"Average LPAR Usage",BOLD|LEFT|VCENTRE|GREEN);
		row++;
		row++;
		
		ds.addLabel(0,row,"Hours",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
		ds.addLabel(0,row,"Reds %",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
		ds.addLabel(0,row,"Yellows %",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
		ds.addLabel(0,row,"Over Cap %",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
		ds.addLabel(0,row,"Max",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
		ds.addLabel(0,row,"Avg",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
		ds.addLabel(0,row,"90% <=",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
		ds.addLabel(0,row,"95% <=",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
		ds.addLabel(0,row,"Last Ent",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
		row++;
		ds.addLabel(0,row,"Date",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
			
		
		for (i=0; i<timeLabels.length; i++) {
			s = timeLabels[i];
			ds.addLabel(0,row+i,s,BOLD|B_ALL_MED);
		}	
				
		
		red = new int[to-from+1];
		yellow = new int[red.length];
		num = new int[red.length];
		uncap = new int[red.length];
		
		for (i=from; i<=to; i++) {			
			for (j=0; j<60*24; j++) {
				f = lparPC[i].getHourData(j);
				
				ent = lparEnt[i].getHourData(j);
				cap = lparStatus[i].getHourCap(j);
				ms = lparStatus[i].getHourMS(j);
				pool = lparStatus[i].getHourPool(j);
				vp = lparVP[i].getHourData(j);
				
				if (f<0 || ent<=0 || ms==null)
					continue;
				
				if (pool.equals("DefaultPool")) {
					free = msCoreConfig[getMSid(ms)].getHourData(j) - msCoreUsed[getMSid(ms)].getHourData(j);
				} else {
					poolID = getProcPoolId(ms,pool);
					free = procPoolConfig[poolID].getHourData(j)-procPoolUsed[poolID].getHourData(j);
				}
				
				if (cap) {
					max = ent;					
				} else {
					if (f+free>=vp)
						max = vp;
					else
						max = f+free;
				}
				map = B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW;
				if (f/max*100>=R_LEVEL) {
					map = map|RED;
					red[i-from]++;
					yellow[i-from]++;
				} else if (f/max*100>=Y_LEVEL) {
					map = map|YELLOW;
					yellow[i-from]++;
				}
				
				num[i-from]++;
				
				if (f>ent) {
					map = map | BOLD;
					uncap[i-from]++;
				}
					
				ds.addFloat(1+i-from,row+j, f, map);

			}
		}
		
		// Add lpar labels and counters
		for (i=from; i<=to; i++) {
			s = lparNames[i];
			if (s.length()>size)
				size=s.length();
			
			map = BOLD|DIAG45|B_ALL_MED;
			if (1d*red[i-from]/num[i-from]>COLOR_LEVEL)
				map = map | RED;
			else if (1d*yellow[i-from]/num[i-from]>COLOR_LEVEL)
				map = map | YELLOW;
			
			ds.addLabel(1+i-from,row-12,s,map);
			
			map = B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW;
			ds.addInteger(1+i-from,row-11, num[i-from], map);
			
			if (num[i-from]==0) {
				map = B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW;
				ds.addFloat(1+i-from,row-10, 0, map);
				ds.addFloat(1+i-from,row-9, 0, map);
				ds.addFloat(1+i-from,row-8, 0, map);
				ds.addFloat(1+i-from,row-7, 0, map);
				ds.addFloat(1+i-from,row-6, 0, map);
				ds.addFloat(1+i-from,row-5, 0, map);
				ds.addFloat(1+i-from,row-4, 0, map);
				ds.addFloat(1+i-from,row-3, 0, map);
			} else {		
				if (1d*red[i-from]/num[i-from]>COLOR_LEVEL)
				map = map | RED;
				ds.addFloat(1+i-from,row-10, 100d*red[i-from]/num[i-from], map);
				
				map = B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW;
				if (1d*yellow[i-from]/num[i-from]>COLOR_LEVEL)
					map = map | YELLOW;
				ds.addFloat(1+i-from,row-9, 100d*yellow[i-from]/num[i-from], map);
				
				map = B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW;
				ds.addFloat(1+i-from,row-8, 100d*uncap[i-from]/num[i-from], map);
				
				col = jxl.CellReferenceHelper.getColumnReference(1+i-from);
				
				ds.addFloat(1+i-from, row-7, lparPC[i].getHourlyMax(), B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW); 
				//ds.addFormula(1+i-from, row-7, "MAX("+col+(row+1)+":"+col+(row+24*60)+")", B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
				
				ds.addFloat(1+i-from, row-6, lparPC[i].getHourlyAvg(), B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
				//ds.addFormula(1+i-from, row-6, "AVERAGE("+col+(row+1)+":"+col+(row+24*60)+")", B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
				
				ds.addFloat(1+i-from, row-5, lparPC[i].getHourly90p(), B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
				//ds.addFormula(1+i-from, row-5, "STDEVP("+col+(row+1)+":"+col+(row+24*60)+")", B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
				
				ds.addFloat(1+i-from, row-4, lparPC[i].getHourly95p(), B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
				//ds.addFormula(1+i-from, row-4, "1282/1000*"+col+(row-5+1)+"+"+col+(row-6+1)+")", B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);

				
				for (j=24*60-1; j>=0; j--)
					if ( lparEnt[i].getHourData(j)>0) break;
				ds.addFloat(1+i-from,row-3, lparEnt[i].getHourData(j), B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);		
			}
		}


		row++;
		
		
		row+=60*24;
		row++;	
	

		ds.setColSize(0, 15);
		for (i=1; i<=(to-from)+1; i++)
			ds.setColSize(i, 7);

		return ds;
	}
	
	
	
	
	private void createSysPoolHourlyUsageSheetExcel(WritableSheet sheet) {
		DataSheet ds = createSysPoolHourlyUsageSheet();
		if (ds==null) return;
		ds.createExcelSheet(sheet);
	}
	
	private void createSysPoolHourlyUsageSheetHTML(String fileName) {
		DataSheet ds = createSysPoolHourlyUsageSheet();
		if (ds==null) return;
		ds.createHTMLSheet(fileName);
		addButton("CPU Pool Hourly",new File(fileName).getName());
	}
	
	private void createSysPoolHourlyUsageSheetCSV(String fileName) {
		DataSheet ds = createSysPoolHourlyUsageSheet();
		if (ds==null) return;
		ds.setSeparator(csvSeparator);
		ds.createCSVSheet(fileName);
	}
	
	private DataSheet createSysPoolHourlyUsageSheet() {
		DataSheet ds = new DataSheet();
		int i,j;
		String s;
		float f,f2;
		int		row;
		int		validMS;
		int		cells;
		int red[], yellow[], num[];
		int map;
		String col;
		
		GregorianCalendar gc=null;
		
		String timeLabels[] = msCoreUsed[0].getHourlyLabels();
		
		
		row=0;	
		
		ds.addLabel(0,row,"Red",BOLD|CENTRE|VCENTRE|RED);
		ds.mergeCells(1, row, 10, row);
		ds.addLabel(1,row,"Usage >= "+R_LEVEL+"% of max allowed size",BOLD);
		row++;
		
		ds.addLabel(0,row,"Yellow",BOLD|CENTRE|VCENTRE|YELLOW);
		ds.mergeCells(1, row, 10, row);
		ds.addLabel(1,row,"Usage >= "+Y_LEVEL+"% of max allowed size",BOLD);
		row++;
		
		ds.addLabel(0,row,"Pool color",BOLD|CENTRE|VCENTRE);
		ds.mergeCells(1, row, 10, row);
		ds.addLabel(1,row,"Colored cells >= "+COLOR_LEVEL*100+"% of total (yellows include reds)",BOLD);
		row++;
		row++;
		
		if (procPoolName!=null)
			cells = managedSystem.length+1+procPoolName.length;
		else
			cells = managedSystem.length+1;
		
		ds.mergeCells(0, row, cells, row);
		ds.addLabel(0,row,"Average Pool Usage",BOLD|CENTRE|VCENTRE|GREEN);
		row++;
				
		if (procPoolName!=null) {
			red = new int[managedSystem.length+procPoolName.length];
			yellow = new int[red.length];
			num = new int[red.length];
		} else {
			red = new int[managedSystem.length];
			yellow = new int[red.length];
			num = new int[red.length];
		}
		
				
		row++;
		
		ds.addLabel(0,row,"Hours",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
		ds.addLabel(0,row,"Reds %",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
		ds.addLabel(0,row,"Yellows %",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
		ds.addLabel(0,row,"Max",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
		ds.addLabel(0,row,"Avg",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
		ds.addLabel(0,row,"90% <=",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
		ds.addLabel(0,row,"95% <=",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
		ds.addLabel(0,row,"Last size",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
		row++;
		ds.addLabel(0,row,"Date",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
		
		for (i=0; i<timeLabels.length; i++) {
			s = timeLabels[i];
			ds.addLabel(0,row+i,s,BOLD|B_ALL_MED);
		}	
		
		
		for (i=0; i<managedSystem.length; i++) {
			
			red[i] = yellow[i] = num[i] = 0;
			
			for (j=0; j<24*60; j++) {
				f = msCoreUsed[i].getHourData(j);
				f2= msCoreConfig[i].getHourData(j);
				
				if (f<0 || f2<=0)
					continue;
				
				f2 = f/f2*100;		// used%
				num[i]++;
				
				map = B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW;  
				if (f2>=R_LEVEL) {
					map = map|RED;
					yellow[i]++;
					red[i]++;
				} else if (f2>=Y_LEVEL) {
					map = map|YELLOW;
					yellow[i]++;
				} 
				
				ds.addFloat(1+i,row+j, f, map);
			}
		}
		
		for (i=0; procPoolName!=null && i<procPoolName.length; i++) {	
			
			red[managedSystem.length+i] = yellow[managedSystem.length+i] = num[managedSystem.length+i] = 0;	
			
			for (j=0; j<24*60; j++) {
				f = procPoolUsed[i].getHourData(j);
				f2= procPoolConfig[i].getHourData(j);
				
				if (f<0 || f2<=0)
					continue;
				
				f2 = f/f2*100;		// used%
				num[managedSystem.length+i]++;
				
				map = B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW;  
				if (f2>=R_LEVEL) {
					map = map|RED;
					yellow[managedSystem.length+i]++;
					red[managedSystem.length+i]++;
				} else if (f2>=Y_LEVEL) {
					map = map|YELLOW;
					yellow[managedSystem.length+i]++;
				} 
				
				ds.addFloat(managedSystem.length+2+i,row+j, f, map);
			}
		}
		
		// Create labels
		for (i=0; i<managedSystem.length; i++) {
			s = managedSystem[i].getVarValues("name")[0];
			
			map = BOLD|DIAG45|B_ALL_MED;
			if (1d*red[i]/num[i]>COLOR_LEVEL)
				map = map | RED;
			else if (1d*yellow[i]/num[i]>COLOR_LEVEL)
				map = map | YELLOW;
			
			ds.addLabel(1+i,row-11,s,map);
		}
		
		for (i=0; procPoolName!=null && i<procPoolName.length; i++) {
			s = procPoolName[i];	
			
			map = BOLD|DIAG45|B_ALL_MED|WRAP;
			if (1d*red[managedSystem.length+i]/num[managedSystem.length+i]>COLOR_LEVEL)
				map = map | RED;
			else if (1d*yellow[managedSystem.length+i]/num[managedSystem.length+i]>COLOR_LEVEL)
				map = map | YELLOW;
			
			ds.addLabel(managedSystem.length+2+i,row-11,s,map);
		}
		
		// Show days, red and yellow counters
		for (i=0; i<managedSystem.length; i++) {
			map = B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW;
			ds.addInteger(1+i,row-10, num[i], map);
			
			if (num[i]==0){
				ds.addInteger(1+i,row-9, 0, map);
				ds.addInteger(1+i,row-8, 0, map);
				ds.addInteger(1+i,row-7, 0, map);
				ds.addInteger(1+i,row-6, 0, map);
				ds.addInteger(1+i,row-5, 0, map);
				ds.addInteger(1+i,row-4, 0, map);
				ds.addInteger(1+i,row-3, 0, map);
			} else {
				if (1d*red[i]/num[i]>COLOR_LEVEL)
					map = map | RED;
				ds.addFloat(1+i,row-9, 100d*red[i]/num[i], map);
				
				map = B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW;
				if (1d*yellow[i]/num[i]>COLOR_LEVEL)
					map = map | YELLOW;
				ds.addFloat(1+i,row-8, 100d*yellow[i]/num[i], map);
				
				col = jxl.CellReferenceHelper.getColumnReference(1+i);
				
				ds.addFloat(1+i, row-7, msCoreUsed[i].getHourlyMax(), B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
				//ds.addFormula(1+i, row-7, "MAX("+col+(row+1)+":"+col+(row+24*60)+")", B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
				
				ds.addFloat(1+i, row-6, msCoreUsed[i].getHourlyAvg(), B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
				//ds.addFormula(1+i, row-6, "AVERAGE("+col+(row+1)+":"+col+(row+24*60)+")", B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
				
				ds.addFloat(1+i, row-5, msCoreUsed[i].getHourly90p(), B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
				//ds.addFormula(1+i, row-5, "STDEVP("+col+(row+1)+":"+col+(row+24*60)+")", B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
				
				ds.addFloat(1+i, row-4, msCoreUsed[i].getHourly95p(), B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
				//ds.addFormula(1+i, row-4, "1282/1000*"+col+(row-5+1)+"+"+col+(row-6+1)+")", B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
				
				

				for (j=24*60-1; j>=0; j--)
					if ( msCoreConfig[i].getHourData(j)>0) break;
				ds.addInteger(1+i,row-3, msCoreConfig[i].getHourData(j), B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
			}
		}
		
		for (i=0; procPoolName!=null && i<procPoolName.length; i++) {
			map = B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW;
			ds.addInteger(managedSystem.length+2+i,row-10, num[managedSystem.length+i], map);
			
			if (num[managedSystem.length+i]==0) {
				ds.addInteger(managedSystem.length+2+i,row-9, 0, map);
				ds.addInteger(managedSystem.length+2+i,row-8, 0, map);
				ds.addInteger(managedSystem.length+2+i,row-7, 0, map);
				ds.addInteger(managedSystem.length+2+i,row-6, 0, map);
				ds.addInteger(managedSystem.length+2+i,row-5, 0, map);
				ds.addInteger(managedSystem.length+2+i,row-4, 0, map);
				ds.addInteger(managedSystem.length+2+i,row-3, 0, map);
			} else {
				if (1d*red[managedSystem.length+i]/num[managedSystem.length+i]>COLOR_LEVEL)
					map = map | RED;
				ds.addFloat(managedSystem.length+2+i,row-9, 100d*red[managedSystem.length+i]/num[managedSystem.length+i], map);
				
				map = B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW;
				if (1d*yellow[managedSystem.length+i]/num[managedSystem.length+i]>COLOR_LEVEL)
					map = map | YELLOW;
				ds.addFloat(managedSystem.length+2+i,row-8, 100d*yellow[managedSystem.length+i]/num[managedSystem.length+i], map);
				
				col = jxl.CellReferenceHelper.getColumnReference(managedSystem.length+2+i);
				
				ds.addFloat(managedSystem.length+2+i, row-7, procPoolUsed[i].getHourlyMax(), B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
				//ds.addFormula(managedSystem.length+2+i, row-7, "MAX("+col+(row+1)+":"+col+(row+24*60)+")", B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
				
				ds.addFloat(managedSystem.length+2+i, row-6, procPoolUsed[i].getHourlyAvg(), B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
				//ds.addFormula(managedSystem.length+2+i, row-6, "AVERAGE("+col+(row+1)+":"+col+(row+24*60)+")", B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
				
				ds.addFloat(managedSystem.length+2+i, row-5, procPoolUsed[i].getHourly90p(), B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
				//ds.addFormula(managedSystem.length+2+i, row-5, "STDEVP("+col+(row+1)+":"+col+(row+24*60)+")", B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
				
				ds.addFloat(managedSystem.length+2+i, row-4, procPoolUsed[i].getHourly95p(), B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
				//ds.addFormula(managedSystem.length+2+i, row-4, "1282/1000*"+col+(row-5+1)+"+"+col+(row-6+1)+")", B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
				
				
				for (j=24*60-1; j>=0; j--)
					if ( procPoolConfig[i].getHourData(j)>0) break;
				ds.addInteger(managedSystem.length+2+i,row-3, procPoolConfig[i].getHourData(j), B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
			}
		}
		row++;
		
		row+=24*60;
		row++;	
		
		ds.setColSize(0, 15);
		for (i=1; i<=managedSystem.length; i++)
			ds.setColSize(i, 8);

		return ds;
	}
	
	
	
	
	
	private void createSysPoolDailyUsageSheetExcel(WritableSheet sheet) {
		DataSheet ds = createSysPoolDailyUsageSheet();
		if (ds!=null)
			ds.createExcelSheet(sheet);
	}
	
	private void createSysPoolDailyUsageSheetHTML(String fileName) {
		DataSheet ds = createSysPoolDailyUsageSheet();
		if (ds!=null) {
			ds.createHTMLSheet(fileName);
			addButton("CPU Pool Daily",new File(fileName).getName());
		}
	}
	
	private void createSysPoolDailyUsageSheetCSV(String fileName) {
		DataSheet ds = createSysPoolDailyUsageSheet();
		if (ds!=null) {
			ds.setSeparator(csvSeparator);
			ds.createCSVSheet(fileName);
		}
	}
	
	private DataSheet createSysPoolDailyUsageSheet() {
		DataSheet ds = new DataSheet();
		int i,j;
		String s;
		float f,f2;
		int		row;
		int		cells;
		int red[], yellow[], num[];
		int map;
		String col;
		
		
		String timeLabels[] = msCoreUsed[0].getDailyLabels();
		
		
		row=0;		
		
		ds.addLabel(0,row,"Red",BOLD|CENTRE|VCENTRE|RED);
		ds.mergeCells(1, row, 10, row);
		ds.addLabel(1,row,"Usage >= "+R_LEVEL+"% of max allowed size",BOLD);
		row++;
		
		ds.addLabel(0,row,"Yellow",BOLD|CENTRE|VCENTRE|YELLOW);
		ds.mergeCells(1, row, 10, row);
		ds.addLabel(1,row,"Usage >= "+Y_LEVEL+"% of max allowed size",BOLD);
		row++;
		
		ds.addLabel(0,row,"Pool color",BOLD|CENTRE|VCENTRE);
		ds.mergeCells(1, row, 10, row);
		ds.addLabel(1,row,"Colored cells >= "+COLOR_LEVEL*100+"% of total (yellows include reds)",BOLD);
		row++;
		row++;
		
		
		if (procPoolName!=null)
			cells = managedSystem.length+1+procPoolName.length;
		else
			cells = managedSystem.length+1;
		
		ds.mergeCells(0, row, cells, row);
		ds.addLabel(0,row,"Average Pool Usage",BOLD|CENTRE|VCENTRE|GREEN);
		row++;
		
		//addLabel(sheet,0,row,"Date",formatLabel(BOLD|VCENTRE|CENTRE|B_ALL_MED));
		
		if (procPoolName!=null) {
			red = new int[managedSystem.length+procPoolName.length];
			yellow = new int[red.length];
			num = new int[red.length];
		} else {
			red = new int[managedSystem.length];
			yellow = new int[red.length];
			num = new int[red.length];
		}

		/*
		for (i=0; i<managedSystem.length; i++) {
			s = managedSystem[i].getVarValues("name")[0];
			//if (s.length()>size)
			//	size=s.length();
			
			addLabel(sheet,1+i,row,s,formatLabel(BOLD|DIAG45|B_ALL_MED));
		}
		
		for (i=0; procPoolName!=null && i<procPoolName.length; i++) {
			s = procPoolName[i];			
			addLabel(sheet,managedSystem.length+2+i,row,s,formatLabel(BOLD|DIAG45|B_ALL_MED|WRAP));
		}
		*/
		
		row++;
		
		ds.addLabel(0,row,"Days",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
		ds.addLabel(0,row,"Reds %",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
		ds.addLabel(0,row,"Yellows %",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
		ds.addLabel(0,row,"Max",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
		ds.addLabel(0,row,"Avg",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
		ds.addLabel(0,row,"90% <=",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
		ds.addLabel(0,row,"95% <=",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
		ds.addLabel(0,row,"Last size",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
		row++;
		ds.addLabel(0,row,"Date",BOLD|VCENTRE|CENTRE|B_ALL_MED);
		row++;
		
		for (i=0; i<timeLabels.length; i++) {
			s = timeLabels[i];
			ds.addLabel(0,row+i,s,BOLD|B_ALL_MED);
		}
		
		
		
		for (i=0; i<managedSystem.length; i++) {	
			
			red[i] = yellow[i] = num[i] = 0;	
			
			for (j=0; j<365; j++) {
				f = msCoreUsed[i].getDayData(j);
				f2 = msCoreConfig[i].getDayData(j);
				
				if (f<0 || f2<=0)
					continue;
				
				f2 = f/f2*100;		// used%
				num[i]++;
				
				map = B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW;  
				if (f2>=R_LEVEL) {
					map = map|RED;
					yellow[i]++;
					red[i]++;
				} else if (f2>=Y_LEVEL) {
					map = map|YELLOW;
					yellow[i]++;
				} 
				
				ds.addFloat(1+i,row+j, f, map);
				
				/*
				if (f>=0)
					addNumber(sheet,1+i,row+j, f, formatFloat(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				*/
			}
		}
		
		for (i=0; procPoolName!=null && i<procPoolName.length; i++) {	
			
			red[managedSystem.length+i] = yellow[managedSystem.length+i] = num[managedSystem.length+i] = 0;	
			
			for (j=0; j<365; j++) {
				f = procPoolUsed[i].getDayData(j);
				f2= procPoolConfig[i].getDayData(j);
				
				if (f<0 || f2<=0)
					continue;
				
				f2 = f/f2*100;		// used%
				num[managedSystem.length+i]++;
				
				map = B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW;  
				if (f2>=R_LEVEL) {
					map = map|RED;
					yellow[managedSystem.length+i]++;
					red[managedSystem.length+i]++;
				} else if (f2>=Y_LEVEL) {
					map = map|YELLOW;
					yellow[managedSystem.length+i]++;
				} 
				
				ds.addFloat(managedSystem.length+2+i,row+j, f, map);
				
				/*
				if (f>=0)
					addNumber(sheet,managedSystem.length+2+i,row+j, f, formatFloat(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				*/
			}
		}
		
		// Create labels	
		for (i=0; i<managedSystem.length; i++) {
			s = managedSystem[i].getVarValues("name")[0];
			
			map = BOLD|DIAG45|B_ALL_MED;
			if (1d*red[i]/num[i]>COLOR_LEVEL)
				map = map | RED;
			else if (1d*yellow[i]/num[i]>COLOR_LEVEL)
				map = map | YELLOW;
			
			ds.addLabel(1+i,row-11,s,map);
		}
		
		for (i=0; procPoolName!=null && i<procPoolName.length; i++) {
			s = procPoolName[i];	
			
			map = BOLD|DIAG45|B_ALL_MED|WRAP;
			if (1d*red[managedSystem.length+i]/num[managedSystem.length+i]>COLOR_LEVEL)
				map = map | RED;
			else if (1d*yellow[managedSystem.length+i]/num[managedSystem.length+i]>COLOR_LEVEL)
				map = map | YELLOW;
			
			ds.addLabel(managedSystem.length+2+i,row-11,s,map);
		}
		
		// Show days, red and yellow counters
		for (i=0; i<managedSystem.length; i++) {
			map = B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW;
			ds.addInteger(1+i,row-10, num[i], map);
			
			if (num[i]==0) {
				ds.addInteger(1+i,row-9, 0, map);
				ds.addInteger(1+i,row-8, 0, map);
				ds.addInteger(1+i,row-7, 0, map);
				ds.addInteger(1+i,row-6, 0, map);
				ds.addInteger(1+i,row-5, 0, map);
				ds.addInteger(1+i,row-4, 0, map);
				ds.addInteger(1+i,row-3, 0, map);
			} else {
				if (1d*red[i]/num[i]>COLOR_LEVEL)
					map = map | RED;
				ds.addFloat(1+i,row-9, 100d*red[i]/num[i], map);
				
				map = B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW;
				if (1d*yellow[i]/num[i]>COLOR_LEVEL)
					map = map | YELLOW;
				ds.addFloat(1+i,row-8, 100d*yellow[i]/num[i], map);
			
				col = jxl.CellReferenceHelper.getColumnReference(1+i);
				
				ds.addFloat(1+i, row-7, msCoreUsed[i].getDailyMax(), B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
				//ds.addFormula(1+i, row-7, "MAX("+col+(row+1)+":"+col+(row+365)+")", B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
						
				ds.addFloat(1+i, row-6, msCoreUsed[i].getDailyAvg(), B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
				//ds.addFormula(1+i, row-6, "AVERAGE("+col+(row+1)+":"+col+(row+365)+")", B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
								
				ds.addFloat(1+i, row-5, msCoreUsed[i].getDaily90p(), B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
				//ds.addFormula(1+i, row-5, "STDEVP("+col+(row+1)+":"+col+(row+365)+")", B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
								
				ds.addFloat(1+i, row-4, msCoreUsed[i].getDaily95p(), B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
				//ds.addFormula(1+i, row-4, "1282/1000*"+col+(row-5+1)+"+"+col+(row-6+1)+")", B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
								
				for (j=364; j>=0; j--)
					if ( msCoreConfig[i].getDayData(j)>0) break;
				ds.addInteger(1+i,row-3, msCoreConfig[i].getDayData(j), B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
			}

 
		}
		
		for (i=0; procPoolName!=null && i<procPoolName.length; i++) {
			map = B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW;
			ds.addInteger(managedSystem.length+2+i,row-10, num[managedSystem.length+i], map);
			
			if (num[managedSystem.length+i]==0) {
				ds.addInteger(managedSystem.length+2+i,row-9, 0, map);
				ds.addInteger(managedSystem.length+2+i,row-8, 0, map);
				ds.addInteger(managedSystem.length+2+i,row-7, 0, map);
				ds.addInteger(managedSystem.length+2+i,row-6, 0, map);
				ds.addInteger(managedSystem.length+2+i,row-5, 0, map);
				ds.addInteger(managedSystem.length+2+i,row-4, 0, map);
				ds.addInteger(managedSystem.length+2+i,row-3, 0, map);
			} else {
				if (1d*red[managedSystem.length+i]/num[managedSystem.length+i]>COLOR_LEVEL)
					map = map | RED;
				ds.addFloat(managedSystem.length+2+i,row-9, 100d*red[managedSystem.length+i]/num[managedSystem.length+i], map);
				
				map = B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW;
				if (1d*yellow[managedSystem.length+i]/num[managedSystem.length+i]>COLOR_LEVEL)
					map = map | YELLOW;
				ds.addFloat(managedSystem.length+2+i,row-8, 100d*yellow[managedSystem.length+i]/num[managedSystem.length+i], map);
			
				col = jxl.CellReferenceHelper.getColumnReference(managedSystem.length+2+i);
				
				ds.addFloat(managedSystem.length+2+i, row-7, procPoolUsed[i].getDailyMax(), B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
				//ds.addFormula(managedSystem.length+2+i, row-7, "MAX("+col+(row+1)+":"+col+(row+365)+")", B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
				
				ds.addFloat(managedSystem.length+2+i, row-6, procPoolUsed[i].getDailyAvg(), B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
				//ds.addFormula(managedSystem.length+2+i, row-6, "AVERAGE("+col+(row+1)+":"+col+(row+365)+")", B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
								
				ds.addFloat(managedSystem.length+2+i, row-5, procPoolUsed[i].getDaily90p(), B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
				//ds.addFormula(managedSystem.length+2+i, row-5, "STDEVP("+col+(row+1)+":"+col+(row+365)+")", B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
				
				ds.addFloat(managedSystem.length+2+i, row-4, procPoolUsed[i].getDaily95p(), B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
				//ds.addFormula(managedSystem.length+2+i, row-4, "1282/1000*"+col+(row-5+1)+"+"+col+(row-6+1)+")", B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
								
				for (j=364; j>=0; j--)
					if ( procPoolConfig[i].getDayData(j)>0) break;
				ds.addInteger(managedSystem.length+2+i,row-3, procPoolConfig[i].getDayData(j), B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW);
			}
		}
				
		row+=365;
		row++;	
	
	
		ds.setColSize(0, 12);
		for (i=1; i<=managedSystem.length; i++)
			ds.setColSize(i, 8);

		return ds;
	}
	
	
	
	
	
	private void createSysPoolUsageSheetExcel(WritableSheet sheet) {
		DataSheet ds = createSysPoolUsageSheet();
		if (ds!=null)
			ds.createExcelSheet(sheet);
	}
	
	private void createSysPoolUsageSheetHTML(String fileName) {
		DataSheet ds = createSysPoolUsageSheet();
		if (ds!=null) {
			ds.createHTMLSheet(fileName);
			addButton("CPU Pool Usage",new File(fileName).getName());
		}
	}
	
	private void createSysPoolUsageSheetCSV(String fileName) {
		DataSheet ds = createSysPoolUsageSheet();
		if (ds!=null) {
			ds.setSeparator(csvSeparator);
			ds.createCSVSheet(fileName);
		}
	}
	
	private DataSheet createSysPoolUsageSheet() {
		DataSheet ds = new DataSheet();
		int i,j;
		String s;
		double f,f2;
		int		size=0;
		int		row;
		int red, yellow;
		int num;
		int map = 0;
		
		
		String timeLabels[] = msCoreUsed[0].getMonthlyLabels();
		
		
		
		row=0;	
		
		ds.addLabel(0,row,"Red",BOLD|CENTRE|VCENTRE|RED);
		ds.mergeCells(1, row, 10, row);
		ds.addLabel(1,row,"Usage >= "+R_LEVEL+"% of max allowed size",BOLD);
		row++;
		
		ds.addLabel(0,row,"Yellow",BOLD|CENTRE|VCENTRE|YELLOW);
		ds.mergeCells(1, row, 10, row);
		ds.addLabel(1,row,"Usage >= "+Y_LEVEL+"% of max allowed size",BOLD);
		row++;
		
		ds.addLabel(0,row,"Pool color",BOLD|CENTRE|VCENTRE);
		ds.mergeCells(1, row, 10, row);
		ds.addLabel(1,row,"Colored cells >= "+COLOR_LEVEL*100+"% of total (yellows include reds)",BOLD);
		row++;
		row++;
		
		ds.mergeCells(0, row, 13, row);
		ds.addLabel(0,row,"Average Pool Usage",BOLD|CENTRE|VCENTRE|GREEN);
		ds.mergeCells(15, row, 28, row);
		ds.addLabel(15,row,"Average Percentage Pool Usage",BOLD|CENTRE|VCENTRE|GREEN);
		row++;
		
		for (i=0; i<timeLabels.length; i++) {
			s = timeLabels[i];
			ds.addLabel(1+i,row,s,BOLD|B_ALL_MED);
			ds.addLabel(16+i,row,s,BOLD|B_ALL_MED);
		}
		

		row++;
		
		
		for (i=0; i<managedSystem.length; i++) {
			s = managedSystem[i].getVarValues("name")[0];
			if (s.length()>size)
				size=s.length();
			
			red = yellow = num = 0;	
			
			for (j=0; j<timeLabels.length; j++) {
				f = msCoreUsed[i].getMonthData(j);
				f2 = msCoreConfig[i].getMonthData(j);
				
				if (f<0 || f2<=0)
					continue;
				
				f2 = f/f2*100;		// used%
				num++;
				
				map = B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW;
				if (f2>=R_LEVEL) {
					map = map|RED;
					yellow++;
					red++;
				} else if (f2>=Y_LEVEL) {
					map = map|YELLOW;
					yellow++;
				} 
				
				ds.addFloat(1+j,row+i, f, map);
				ds.addFloat(16+j,row+i, f2, map);
				
				/*
				if (f>=0)
					addNumber(sheet,1+j,row+i, f, formatFloat(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				if (f>=0 && f2>0)
					addNumber(sheet,16+j,row+i, f/f2*100, formatFloat(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				*/
			}
			
			map = BOLD|VCENTRE|CENTRE|B_ALL_MED;
			
			if (1d*red/num>COLOR_LEVEL)
				map = map | RED;
			else if (1d*yellow/num>COLOR_LEVEL)
				map = map | YELLOW;

			ds.addLabel(0,row+i,s,map);
			ds.addLabel(15,row+i,s,map);
		}
		row+=i;
		row++;
		
		
		ds.mergeCells(0, row, 13, row);
		ds.addLabel(0,row,"Peak Pool Usage",BOLD|CENTRE|VCENTRE|GREEN);
		ds.mergeCells(15, row, 28, row);
		ds.addLabel(15,row,"Peak Percentage Pool Usage",BOLD|CENTRE|VCENTRE|GREEN);
		row++;
		
		for (i=0; i<timeLabels.length; i++) {
			s = timeLabels[i];
			ds.addLabel(1+i,row,s,BOLD|B_ALL_MED);
			ds.addLabel(16+i,row,s,BOLD|B_ALL_MED);
		}
		

		row++;
		
		
		for (i=0; i<managedSystem.length; i++) {
			s = managedSystem[i].getVarValues("name")[0];
			if (s.length()>size)
				size=s.length();
			
			red = yellow = num = 0;	
			
			for (j=0; j<timeLabels.length; j++) {
				f = msCoreUsed[i].getMonthlyMax(j);
				f2 = msCoreConfig[i].getMonthlyMax(j);
				
				if (f<0 || f2<=0)
					continue;
				
				f2 = f/f2*100;		// used%
				num++;
				
				map = B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW;
				if (f2>=R_LEVEL) {
					map = map|RED;
					yellow++;
					red++;
				} else if (f2>=Y_LEVEL) {
					map = map|YELLOW;
					yellow++;
				} 
				
				ds.addFloat(1+j,row+i, f, map);
				ds.addFloat(16+j,row+i, f2, map);
				
				/*
				if (f>=0)
					addNumber(sheet,1+j,row+i, f, formatFloat(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				if (f>=0 && f2>0)
					addNumber(sheet,16+j,row+i, f/f2*100, formatFloat(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				*/
			}
			
			map = BOLD|VCENTRE|CENTRE|B_ALL_MED;
			
			if (1d*red/num>COLOR_LEVEL)
				map = map | RED;
			else if (1d*yellow/num>COLOR_LEVEL)
				map = map | YELLOW;

			ds.addLabel(0,row+i,s,map);
			ds.addLabel(15,row+i,s,map);
		}
		row+=i;
		row++;
		
		ds.mergeCells(0, row, 13, row);
		ds.addLabel(0,row,"Average Pool Size",BOLD|CENTRE|VCENTRE|GREEN);
		row++;
		
		for (i=0; i<timeLabels.length; i++) {
			s = timeLabels[i];
			ds.addLabel(1+i,row,s,BOLD|B_ALL_MED);
		}
		

		row++;
		
		
		for (i=0; i<managedSystem.length; i++) {
			s = managedSystem[i].getVarValues("name")[0];
			if (s.length()>size)
				size=s.length();
			
			ds.addLabel(0,row+i,s,BOLD|VCENTRE|CENTRE|B_ALL_MED);
			
			for (j=0; j<timeLabels.length; j++) {
				f = msCoreConfig[i].getMonthlyAvg(j); 
				if (f>=0)
					ds.addFloat(1+j,row+i, f, B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW);			
			}
		}
		row+=i;
		row++;
		
		
		ds.mergeCells(0, row, 13, row);
		ds.addLabel(0,row,"Average Pool Unassigned",BOLD|CENTRE|VCENTRE|GREEN);
		row++;
		
		for (i=0; i<timeLabels.length; i++) {
			s = timeLabels[i];
			ds.addLabel(1+i,row,s,BOLD|B_ALL_MED);
		}
		

		row++;
		
		
		for (i=0; i<managedSystem.length; i++) {
			s = managedSystem[i].getVarValues("name")[0];
			if (s.length()>size)
				size=s.length();
			
			ds.addLabel(0,row+i,s,BOLD|VCENTRE|CENTRE|B_ALL_MED);
			
			for (j=0; j<timeLabels.length; j++) {
				f = msCoreAvail[i].getMonthlyAvg(j);
				if (f>=0)
					ds.addFloat(1+j,row+i, f, B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW);			
			}
		}
		row+=i;
		row++;
		
		
		ds.mergeCells(0, row, 13, row);
		ds.addLabel(0,row,"Average Virtual Pool Usage",BOLD|CENTRE|VCENTRE|GREEN);
		ds.mergeCells(15, row, 28, row);
		ds.addLabel(15,row,"Average Percentage Virtual Pool Usage",BOLD|CENTRE|VCENTRE|GREEN);
		row++;
		
		
		for (i=0; i<timeLabels.length; i++) {
			s = timeLabels[i];
			ds.addLabel(1+i,row,s,BOLD|B_ALL_MED);
			ds.addLabel(16+i,row,s,BOLD|B_ALL_MED);
		}
		
		row++;
		
		
		for (i=0; procPoolName!=null && i<procPoolName.length; i++) {
			s = procPoolName[i];
			
			red = yellow = num = 0;	
			
			for (j=0; j<timeLabels.length; j++) {
				f = procPoolUsed[i].getMonthlyAvg(j);
				f2 = procPoolConfig[i].getMonthlyAvg(j);
				
				if (f<0 || f2<=0)
					continue;
				
				f2 = f/f2*100;		// used%
				num++;
				
				map = B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW;
				if (f2>=R_LEVEL) {
					map = map|RED;
					yellow++;
					red++;
				} else if (f2>=Y_LEVEL) {
					map = map|YELLOW;
					yellow++;
				} 
				
				ds.addFloat(1+j,row+i, f, map);
				ds.addFloat(16+j,row+i, f2, map);
				
				/*
				if (f>=0)
					addNumber(sheet,1+j,row+i, f, formatFloat(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				if (f>=0 && f2>0)
					addNumber(sheet,16+j,row+i, f/f2*100, formatFloat(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				*/
			}
			
			map = BOLD|VCENTRE|CENTRE|B_ALL_MED|WRAP;
			
			if (1d*red/num>COLOR_LEVEL)
				map = map | RED;
			else if (1d*yellow/num>COLOR_LEVEL)
				map = map | YELLOW;

			ds.addLabel(0,row+i,s,map);
			ds.addLabel(15,row+i,s,map);						
		}
		row+=i;
		row++;
		
		
		ds.mergeCells(0, row, 13, row);
		ds.addLabel(0,row,"Peak Virtual Pool Usage",BOLD|CENTRE|VCENTRE|GREEN);
		ds.mergeCells(15, row, 28, row);
		ds.addLabel(15,row,"Peak Percentage Virtual Pool Usage",BOLD|CENTRE|VCENTRE|GREEN);
		row++;
		
		for (i=0; i<timeLabels.length; i++) {
			s = timeLabels[i];
			ds.addLabel(1+i,row,s,BOLD|B_ALL_MED);
			ds.addLabel(16+i,row,s,BOLD|B_ALL_MED);
		}
		

		row++;
		
		
		for (i=0; procPoolName!=null && i<procPoolName.length; i++) {
			s = procPoolName[i];
			
			red = yellow = num = 0;
			
			for (j=0; j<timeLabels.length; j++) {
				f = procPoolUsed[i].getMonthlyMax(j);
				f2 = procPoolConfig[i].getMonthlyMax(j);
				
				if (f<0 || f2<=0)
					continue;
				
				f2 = f/f2*100;		// used%
				num++;
				
				map = B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW;
				if (f2>=R_LEVEL) {
					map = map|RED;
					yellow++;
					red++;
				} else if (f2>=Y_LEVEL) {
					map = map|YELLOW;
					yellow++;
				}
				
				ds.addFloat(1+j,row+i, f, map);
				ds.addFloat(16+j,row+i, f2, map);
			}
			
			map = BOLD|VCENTRE|CENTRE|B_ALL_MED|WRAP;
			
			if (1d*red/num>COLOR_LEVEL)
				map = map | RED;
			else if (1d*yellow/num>COLOR_LEVEL)
				map = map | YELLOW;

			ds.addLabel(0,row+i,s,map);
			ds.addLabel(15,row+i,s,map);	
		}
		row+=i;
		row++;
		
		
		ds.mergeCells(0, row, 13, row);
		ds.addLabel(0,row,"Average Virtual Pool Size",BOLD|CENTRE|VCENTRE|GREEN);
		row++;
		
		for (i=0; i<timeLabels.length; i++) {
			s = timeLabels[i];
			ds.addLabel(1+i,row,s,BOLD|B_ALL_MED);
		}
		
		
		row++;
		
		
		for (i=0; procPoolName!=null && i<procPoolName.length; i++) {
			s = procPoolName[i];
			
			ds.addLabel(0,row+i,s,BOLD|VCENTRE|CENTRE|B_ALL_MED|WRAP);
			
			for (j=0; j<timeLabels.length; j++) {
				f = procPoolConfig[i].getMonthlyMax(j);
				if (f>=0)
					ds.addFloat(1+j,row+i, f, B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW);
			}
		}
		row+=i;
		row++;
		
	
		ds.setColSize(0, size+6);
		ds.setColSize(15, size+6);

		return ds;
	}
	
	
	
	
	
	
	
	
	
	
	private int getMSid(String ms) {
		int i;
		String name;
		
		for (i=0; i<managedSystem.length; i++) {
			name = managedSystem[i].getVarValues("name")[0];
			if (ms.equals(name))
				return i;
		}
		
		return -1;
	}
	
	
	private void createLparCoreUsageSheetExcel(WritableSheet sheet) {
		DataSheet ds = createLparCoreUsageSheet();
		if (ds!=null)
			ds.createExcelSheet(sheet);
	}
	
	private void createLparCoreUsageSheetHTML(String fileName) {
		DataSheet ds = createLparCoreUsageSheet();
		if (ds!=null) {
			ds.createHTMLSheet(fileName);
			addButton("LPAR CPU Usage",new File(fileName).getName());
		}
	}
	
	private void createLparCoreUsageSheetCSV(String fileName) {
		DataSheet ds = createLparCoreUsageSheet();
		if (ds!=null) {
			ds.setSeparator(csvSeparator);
			ds.createCSVSheet(fileName);
		}
	}
	
	private DataSheet createLparCoreUsageSheet() {
		DataSheet ds = new DataSheet();
		int i,j;
		String s;
		double f,f2;
		int		size=0;
		int		row;
		boolean cap;
		String	ms, pool;
		double	vp;
		double	free;
		int		poolID;
		double max;
		int map=0;
		
		int red, yellow, num;
		
		
		String timeLabels[] = msCoreUsed[0].getMonthlyLabels();
		
		
		for (i=0; i<lparNames.length; i++) 
			if (lparNames[i].length()>size)
				size=lparNames[i].length();
		
	
		
		row = 0;
		
		ds.addLabel(0,row,"Red",BOLD|CENTRE|VCENTRE|RED);
		ds.mergeCells(1, row, 10, row);
		ds.addLabel(1,row,"Usage >= "+R_LEVEL+"% of max allowed size",BOLD);
		row++;
		
		ds.addLabel(0,row,"Yellow",BOLD|CENTRE|VCENTRE|YELLOW);
		ds.mergeCells(1, row, 10, row);
		ds.addLabel(1,row,"Usage >= "+Y_LEVEL+"% of max allowed size",BOLD);
		row++;
		
		ds.addLabel(0,row,"Bold",BOLD|CENTRE|VCENTRE);
		ds.mergeCells(1, row, 10, row);
		ds.addLabel(1,row,"Usage > entitled capacity",BOLD);
		row++;
		
		ds.addLabel(0,row,"LPAR color",BOLD|CENTRE|VCENTRE);
		ds.mergeCells(1, row, 10, row);
		ds.addLabel(1,row,"Colored cells >= "+COLOR_LEVEL*100+"% of total (yellows include reds)",BOLD);
		row++;
		row++;
		
		ds.mergeCells(0, row, 13, row);
		ds.addLabel(0,row,"Average Processor Usage",BOLD|CENTRE|VCENTRE|GREEN);
		ds.mergeCells(15, row, 28, row);
		ds.addLabel(15,row,"Average Entitlement Percentage Usage",BOLD|CENTRE|VCENTRE|GREEN);
		row++;
		
		for (i=0; i<timeLabels.length; i++) {
			s = timeLabels[i];
			ds.addLabel(1+i,row,s,BOLD|B_ALL_MED);
			ds.addLabel(16+i,row,s,BOLD|B_ALL_MED);
		}


		row++;
		
		
		for (i=0; i<lparNames.length; i++) {
			
			red = yellow = num = 0;
			
			for (j=0; j<timeLabels.length; j++) {
				f = lparPC[i].getMonthlyAvg(j);
				f2 = lparEnt[i].getMonthlyAvg(j);
				cap = lparStatus[i].getMonthCap(j);
				ms = lparStatus[i].getMonthMS(j);
				pool = lparStatus[i].getMonthPool(j);
				vp = lparVP[i].getMonthlyAvg(j);
				
				if (f<0 || f2<=0 || pool==null || ms==null || vp<0 )
					continue;
				
				if (pool.equals("DefaultPool")) {
					free = msCoreConfig[getMSid(ms)].getMonthlyAvg(j) - msCoreUsed[getMSid(ms)].getMonthlyAvg(j);
				} else {
					poolID = getProcPoolId(ms,pool);
					free = procPoolConfig[poolID].getMonthlyAvg(j)-procPoolUsed[poolID].getMonthlyAvg(j);
				}
				
				if (cap) {
					max = f2;					
				} else {
					if (f+free>=vp)
						max = vp;
					else
						max = f+free;
				}
				map = B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW;
				if (f/max*100>=R_LEVEL) {
					map = map|RED;
					yellow++;
					red++;
				} else if (f/max*100>=Y_LEVEL) {
					map = map|YELLOW;
					yellow++;
				}
				
				if (f>f2)
					map = map | BOLD;
					
				ds.addFloat(1+j,row+i, f, map);
				ds.addFloat(16+j,row+i, f/f2*100, map);
				
				num++;
			}
			
			map = BOLD|VCENTRE|CENTRE|B_ALL_MED;
			
			if (1d*red/num>COLOR_LEVEL)
				map = map | RED;
			else if (1d*yellow/num>COLOR_LEVEL)
				map = map | YELLOW;

			ds.addLabel(0,row+i,lparNames[i],map);
			ds.addLabel(15,row+i,lparNames[i],map);
		}
		row+=i;
		row++;
		
		
		
		ds.mergeCells(0, row, 13, row);
		ds.addLabel(0,row,"Peak Processor Usage",BOLD|CENTRE|VCENTRE|GREEN);
		ds.mergeCells(15, row, 28, row);
		ds.addLabel(15,row,"Peak Entitlement Percentage Usage",BOLD|CENTRE|VCENTRE|GREEN);
		row++;
		
		for (i=0; i<timeLabels.length; i++) {
			s = timeLabels[i];
			ds.addLabel(1+i,row,s,BOLD|B_ALL_MED);
			ds.addLabel(16+i,row,s,BOLD|B_ALL_MED);
		}
		

		row++;
		
		
		for (i=0; i<lparNames.length; i++) {
			
			ds.addLabel(0,row+i,lparNames[i],BOLD|VCENTRE|CENTRE|B_ALL_MED);
			ds.addLabel(15,row+i,lparNames[i],BOLD|VCENTRE|CENTRE|B_ALL_MED);
			
			for (j=0; j<timeLabels.length; j++) {
				f = lparPC[i].getMonthlyMax(j);
				f2 = lparEnt[i].getMonthlyMax(j);
				
				if (f>=0)
					ds.addFloat(1+j,row+i, f, B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW);
				if (f>=0 && f2>0)
					ds.addFloat(16+j,row+i, f/f2*100, B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW);
			}
		}
		row+=i;
		row++;
		
		
		
		ds.mergeCells(0, row, 13, row);
		ds.addLabel(0,row,"Average Entitlement",BOLD|CENTRE|VCENTRE|GREEN);
		ds.mergeCells(15, row, 28, row);
		ds.addLabel(15,row,"Average VPs",BOLD|CENTRE|VCENTRE|GREEN);
		row++;
		
		for (i=0; i<timeLabels.length; i++) {
			s = timeLabels[i];
			ds.addLabel(1+i,row,s,BOLD|B_ALL_MED);
			ds.addLabel(16+i,row,s,BOLD|B_ALL_MED);
		}
		
		
		row++;
		
		
		for (i=0; i<lparNames.length; i++) {
			
			ds.addLabel(0,row+i,lparNames[i],BOLD|VCENTRE|CENTRE|B_ALL_MED);
			
			for (j=0; j<timeLabels.length; j++) {
				f = lparEnt[i].getMonthlyAvg(j);
				if (f>=0)
					ds.addFloat(1+j,row+i, f, B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW);
				f = lparVP[i].getMonthlyAvg(j);
				if (f>=0)
					ds.addFloat(16+j,row+i, f, B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW);
			}
		}
		row+=i;
		row++;
		

		ds.setColSize(0, size+6);
		ds.setColSize(15, size+6);

		return ds;
	}
	
	
	
	
	
	private void createSystemsSheetExcel(WritableSheet sheet) {
		DataSheet ds = createSystemsSheetRowBased();
		if (ds!=null)
			ds.createExcelSheet(sheet);
	}
	
	private void createSystemsSheetHTML(String fileName) {
		DataSheet ds = createSystemsSheetRowBased();
		if (ds!=null) {
			ds.createHTMLSheet(fileName);
			addButton("Systems",new File(fileName).getName());
		}
	}
	
	private void createSystemsSheetCSV(String fileName) {
		DataSheet ds = createSystemsSheetRowBased();
		if (ds!=null) {
			ds.setSeparator(csvSeparator);
			ds.createCSVSheet(fileName);
		}
	}
	
	private DataSheet createSystemsSheetRowBased() {
		int i,j;
		GenericData gd[];
		GenericData lpar[];
		int row,col;
		int nameSize=0;
		String s;
		int vp=0;
		int ded=0;
		int numLpar=0;
		double d;
		
		DataSheet ds = new DataSheet();
			
		row = 0;
		
		/*
		 * Setup titles
		 */ 
		
		ds.addLabel(0,0,"Managed System",BOLD|VCENTRE|B_ALL_MED|GREEN);
		ds.addLabel(1,0,"Status",BOLD|VCENTRE|B_ALL_MED|GREEN);
		ds.addLabel(2,0,"Type Model",BOLD|B_ALL_MED|WRAP|GREEN);
		ds.addLabel(3,0,"Serial",BOLD|B_ALL_MED|GREEN);
		ds.addLabel(4,0,"GHz",BOLD|B_ALL_MED|GREEN);
		ds.addLabel(5,0,"CPU Type",BOLD|B_ALL_MED|GREEN);
		
		
		ds.addLabel(6,0,"Tot Cores",BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_LOW|WRAP|GREEN);
		ds.addLabel(7,0,"Act Cores",BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_LOW|WRAP|GREEN);
		ds.addLabel(8,0,"Deconf Cores",BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_LOW|WRAP|GREEN);
		ds.addLabel(9,0,"Curr Avail Cores",BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_LOW|WRAP|GREEN);
		ds.addLabel(10,0,"Pend Avail Cores",BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_MED|WRAP|GREEN);
		
		ds.addLabel(11,0,"Ded Cores",BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_LOW|WRAP|GREEN);
		ds.addLabel(12,0,"Pool Size",BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_LOW|WRAP|GREEN);
		ds.addLabel(13,0,"Virt Procs",BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_MED|WRAP|GREEN);
		
		ds.addLabel(14,0,"#LPAR",BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_MED|WRAP|GREEN);
		
		ds.addLabel(15,0,"Tot GB",BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_LOW|WRAP|GREEN);
		ds.addLabel(16,0,"Act GB",BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_LOW|WRAP|GREEN);
		ds.addLabel(17,0,"Deconf GB",BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_LOW|WRAP|GREEN);
		ds.addLabel(18,0,"Firm GB",BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_LOW|WRAP|GREEN);
		ds.addLabel(19,0,"Curr Avail GB",BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_LOW|WRAP|GREEN);
		ds.addLabel(20,0,"Pend Avail GB",BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_MED|WRAP|GREEN);
	
		ds.addLabel(21,0,"Perf Sample Rate",BOLD|B_ALL_MED|WRAP|GREEN);
		
		ds.addLabel(22,0,"Mgr #1",BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_LOW|GREEN);
		ds.addLabel(23,0,"Mgr #2",BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_MED|GREEN);
		
		ds.addLabel(24,0,"Prim SP",BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_LOW|WRAP|GREEN);
		ds.addLabel(25,0,"Sec SP",BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_MED|WRAP|GREEN);
	
		ds.addLabel(26,0,"EC Number",BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_LOW|WRAP|GREEN);
		ds.addLabel(27,0,"IPL Level",BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_LOW|WRAP|GREEN);
		ds.addLabel(28,0,"Activated Level",BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_LOW|WRAP|GREEN);
		ds.addLabel(29,0,"Deferred Level",BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_MED|WRAP|GREEN);

		
		row =1;
		
		
		/*
		 * Show values, each system on one line
		 */
		for (i=0; i<managedSystem.length; i++) {
			s = managedSystem[i].getVarValues("name")[0];
			if (s.length() > nameSize)
				nameSize = s.length();
			
			vp=ded=numLpar=0;
			lpar = managedSystem[i].getObjects(PROC_LPAR);
			for (j=0; lpar!=null && j<lpar.length; j++) {
				numLpar++;
				if (lpar[j].getVarValues("curr_proc_mode")[0].equals("shared")) {
					vp += textToInt(lpar[j].getVarValues("run_procs")[0]);
				} else
					ded += textToInt(lpar[j].getVarValues("run_procs")[0]);
			}
			
			col=0;
			
			
			ds.addLabel(col++,row,s,BOLD);
			if (managerType == M_SDMC || managerType == M_FSM)
				ds.addLabel(col++,row,managedSystem[i].getVarValues("primary_state")[0],B_ALL_LOW);
			else
				ds.addLabel(col++,row,managedSystem[i].getVarValues("state")[0],B_ALL_LOW|RIGHT);
			ds.addLabel(col++,row,managedSystem[i].getVarValues("type_model")[0],B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW|RIGHT);
			ds.addLabel(col++,row,managedSystem[i].getVarValues("serial_num")[0],B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW|RIGHT);
			if (managedSystem[i].getVarValues("frequency") != null)
				ds.addFloat(col++,row,managedSystem[i].getVarValues("frequency"),0,B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW|RIGHT);
			else
				col++;
			if (managedSystem[i].getVarValues("cpu_type") != null)
				ds.addLabel(col++,row,managedSystem[i].getVarValues("cpu_type"),0,B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW|RIGHT);
			else
				col++;
			
			gd = managedSystem[i].getObjects(PROC);
			if (gd!=null && gd[0]!=null) {		
				ds.addInteger(col++,row,gd[0].getVarValues("installed_sys_proc_units"),0,B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW);
				ds.addInteger(col++,row,gd[0].getVarValues("configurable_sys_proc_units"),0,B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW);
				ds.addInteger(col++,row,gd[0].getVarValues("deconfig_sys_proc_units"),0,B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW);
				ds.addFloat(col++,row, gd[0].getVarValues("curr_avail_sys_proc_units"),0, B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW);
				ds.addFloat(col++,row, gd[0].getVarValues("pend_avail_sys_proc_units"),0, B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW);
				ds.addInteger(col++,row, ded, B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW);
				
				float pool;
				try {
					pool = Float.parseFloat(gd[0].getVarValues("configurable_sys_proc_units")[0]) - 
							Float.parseFloat(gd[0].getVarValues("deconfig_sys_proc_units")[0]) -
							ded;
					ds.addInteger(col++, row, pool, B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW);
				} catch (NumberFormatException nfe) {
					ds.addLabel(col++,row,"NaN",B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW|RIGHT);
				}
			} else {	
				col = col + 5;
				ds.addInteger(col++,row, ded, B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW);
				col++;
			}
			
			
			ds.addInteger(col++,row, vp, B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW);
			
			ds.addInteger(col++,row, numLpar, B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW);
			
			gd = managedSystem[i].getObjects(MEM);
			if (gd!=null && gd[0]!=null) {	
				ds.addFloatDiv1024(col++, row, gd[0].getVarValues("installed_sys_mem"), 0, B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW);
				ds.addFloatDiv1024(col++, row, gd[0].getVarValues("configurable_sys_mem"), 0, B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW);
				ds.addFloatDiv1024(col++, row, gd[0].getVarValues("deconfig_sys_mem"), 0, B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW);
				ds.addFloatDiv1024(col++, row, gd[0].getVarValues("sys_firmware_mem"), 0, B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW);
				ds.addFloatDiv1024(col++, row, gd[0].getVarValues("curr_avail_sys_mem"), 0, B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW);
				ds.addFloatDiv1024(col++, row, gd[0].getVarValues("pend_avail_sys_mem"), 0, B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW);
			} else
				col = col + 6;
			
			ds.addInteger( col++, row, managedSystem[i].getVarValues("sample_rate"),0, B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW);
			
			ds.addLabel( col++, row, managedSystem[i].getVarValues("HMC1_name"),0, B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW|RIGHT);
			ds.addLabel( col++, row, managedSystem[i].getVarValues("HMC2_name"),0, B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW|RIGHT);
			
			ds.addLabel( col++, row, managedSystem[i].getVarValues("ipaddr"),0, B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW|RIGHT);
			ds.addLabel( col++, row, managedSystem[i].getVarValues("ipaddr_secondary"),0, B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW|RIGHT);

			
			gd = managedSystem[i].getObjects(SYSPOWERLIC);
			if (gd!=null && gd[0]!=null) {	
				ds.addLabel( col++, row, gd[0].getVarValues("ecnumber"),0, B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW|RIGHT);
				ds.addInteger( col++, row, gd[0].getVarValues("platform_ipl_level"),0, B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW);
				ds.addInteger( col++, row, gd[0].getVarValues("activated_level"),0, B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW);
				if (gd[0].getVarValues("deferred_level")[0].equalsIgnoreCase("None"))
					ds.addLabel( col++, row, gd[0].getVarValues("deferred_level"),0, B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW|RIGHT);
				else
					ds.addInteger( col++, row, gd[0].getVarValues("deferred_level"),0, B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW);
			} else
				col = col + 4;
			
			
			row++;
		}	
		
		ds.setColSize(0, nameSize+6);
		
		ds.setColSize(1, 9);
		ds.setColSize(2, 9);
		ds.setColSize(3, 9);
		ds.setColSize(4, 9);
		ds.setColSize(5, 20);
		
		ds.setColSize(6, 6);
		ds.setColSize(7, 6);
		ds.setColSize(8, 6);
		
		ds.setColSize(9, 9);
		ds.setColSize(10, 9);
		
		ds.setColSize(11, 7);
		ds.setColSize(12, 7);
		ds.setColSize(13, 7);
		
		ds.setColSize(14, 7);
		
		ds.setColSize(15, 9);	
		ds.setColSize(16, 9);
		ds.setColSize(17, 9);
		ds.setColSize(18, 9);
		ds.setColSize(19, 9);
		ds.setColSize(20, 9);
		ds.setColSize(21, 9);
		
		ds.setColSize(22, 15);
		ds.setColSize(23, 15);	
		ds.setColSize(24, 15);
		ds.setColSize(25, 15);
	
		ds.setColSize(26, 9);
		ds.setColSize(27, 9);
		ds.setColSize(28, 9);
		ds.setColSize(29, 9);
		
		return ds;	
	}

	
	
	private void createSystemsSheetRowBased(WritableSheet sheet) throws RowsExceededException, WriteException {
		int i,j;
		GenericData gd[];
		GenericData lpar[];
		int row, col;
		int nameSize=0;
		String s;
		int vp=0;
		int ded=0;
		int numLpar=0;
		double d;
			
		row = 0;
		col = 0;
		
		/*
		 * Setup titles
		 */ 
		
		addLabel(sheet,0,0,"Managed System",formatLabel(BOLD|VCENTRE|B_ALL_MED|GREEN));
		addLabel(sheet,1,0,"Status",formatLabel(BOLD|VCENTRE|B_ALL_MED|GREEN));
		addLabel(sheet,2,0,"Type Model",formatLabel(BOLD|B_ALL_MED|WRAP|GREEN));
		addLabel(sheet,3,0,"Serial",formatLabel(BOLD|B_ALL_MED|GREEN));
		addLabel(sheet,4,0,"Freq",formatLabel(BOLD|B_ALL_MED|GREEN));
		
		addLabel(sheet,5,0,"Tot Cores",formatLabel(BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_LOW|WRAP|GREEN));
		addLabel(sheet,6,0,"Act Cores",formatLabel(BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_LOW|WRAP|GREEN));
		addLabel(sheet,7,0,"Deconf Cores",formatLabel(BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_LOW|WRAP|GREEN));
		addLabel(sheet,8,0,"Curr Avail Cores",formatLabel(BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_LOW|WRAP|GREEN));
		addLabel(sheet,9,0,"Pend Avail Cores",formatLabel(BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_MED|WRAP|GREEN));
		
		addLabel(sheet,10,0,"Ded Cores",formatLabel(BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_LOW|WRAP|GREEN));
		addLabel(sheet,11,0,"Pool Size",formatLabel(BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_LOW|WRAP|GREEN));
		addLabel(sheet,12,0,"Virt Procs",formatLabel(BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_MED|WRAP|GREEN));
		
		addLabel(sheet,13,0,"#LPAR",formatLabel(BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_MED|WRAP|GREEN));
		
		addLabel(sheet,14,0,"Tot GB",formatLabel(BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_LOW|WRAP|GREEN));
		addLabel(sheet,15,0,"Act GB",formatLabel(BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_LOW|WRAP|GREEN));
		addLabel(sheet,16,0,"Deconf GB",formatLabel(BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_LOW|WRAP|GREEN));
		addLabel(sheet,17,0,"Firm GB",formatLabel(BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_LOW|WRAP|GREEN));
		addLabel(sheet,18,0,"Curr Avail GB",formatLabel(BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_LOW|WRAP|GREEN));
		addLabel(sheet,19,0,"Pend Avail GB",formatLabel(BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_MED|WRAP|GREEN));
	
		addLabel(sheet,20,0,"Perf Sample Rate",formatLabel(BOLD|B_ALL_MED|WRAP|GREEN));
		
		addLabel(sheet,21,0,"Mgr #1",formatLabel(BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_LOW|GREEN));
		addLabel(sheet,22,0,"Mgr #2",formatLabel(BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_MED|GREEN));
		
		addLabel(sheet,23,0,"Prim SP",formatLabel(BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_LOW|WRAP|GREEN));
		addLabel(sheet,24,0,"Sec SP",formatLabel(BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_MED|WRAP|GREEN));
	
		addLabel(sheet,25,0,"EC Number",formatLabel(BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_LOW|WRAP|GREEN));
		addLabel(sheet,26,0,"IPL Level",formatLabel(BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_LOW|WRAP|GREEN));
		addLabel(sheet,27,0,"Activated Level",formatLabel(BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_LOW|WRAP|GREEN));
		addLabel(sheet,28,0,"Deferred Level",formatLabel(BOLD|B_BOTTOM_MED|B_TOP_MED|B_RIGHT_MED|WRAP|GREEN));

		
		row =1;
		
		
		/*
		 * Show values, each system on one line
		 */
		for (i=0; i<managedSystem.length; i++) {
			s = managedSystem[i].getVarValues("name")[0];
			if (s.length() > nameSize)
				nameSize = s.length();
			
			vp=ded=numLpar=0;
			lpar = managedSystem[i].getObjects(PROC_LPAR);
			for (j=0; lpar!=null && j<lpar.length; j++) {
				numLpar++;
				if (lpar[j].getVarValues("curr_proc_mode")[0].equals("shared")) {
					vp += textToInt(lpar[j].getVarValues("run_procs")[0]);
				} else
					ded += textToInt(lpar[j].getVarValues("run_procs")[0]);
			}
			
			col=0;			
			
			addLabel(sheet,col++,row,s,formatLabel(BOLD));
			if (managerType == M_SDMC || managerType == M_FSM)
				addLabel(sheet,col++,row,managedSystem[i].getVarValues("primary_state")[0],formatLabel(B_ALL_LOW));
			else
				addLabel(sheet,col++,row,managedSystem[i].getVarValues("state")[0],formatLabel(B_ALL_LOW|RIGHT));
			addLabel(sheet,col++,row,managedSystem[i].getVarValues("type_model")[0],formatLabel(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW|RIGHT));
			addLabel(sheet,col++,row,managedSystem[i].getVarValues("serial_num")[0],formatLabel(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW|RIGHT));
			addLabel(sheet,col++,row,managedSystem[i].getVarValues("frequency")[0],formatLabel(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW|RIGHT));
			
			gd = managedSystem[i].getObjects(PROC);
			if (gd!=null && gd[0]!=null) {		
				addNumber(sheet,col++,row,gd[0].getVarValues("installed_sys_proc_units"),0,formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				addNumber(sheet,col++,row,gd[0].getVarValues("configurable_sys_proc_units"),0,formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				addNumber(sheet,col++,row,gd[0].getVarValues("deconfig_sys_proc_units"),0,formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				addNumber(sheet,col++,row, gd[0].getVarValues("curr_avail_sys_proc_units"),0, formatFloat(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				addNumber(sheet,col++,row, gd[0].getVarValues("pend_avail_sys_proc_units"),0, formatFloat(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
			}
			
			addNumber(sheet,col++,row, ded, formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
			Formula f = new Formula(col++, row, "F"+(row+1)+"-G"+(row+1)+"-J"+(row+1),formatInt(B_RIGHT_LOW|B_BOTTOM_LOW|B_LEFT_LOW));
			sheet.addCell(f); 
			addNumber(sheet,col++,row, vp, formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
			
			addNumber(sheet,col++,row, numLpar, formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
			
			gd = managedSystem[i].getObjects(MEM);
			if (gd!=null && gd[0]!=null) {	
				addNumberDiv1024(sheet, col++, row, gd[0].getVarValues("installed_sys_mem"), 0, formatFloat(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				addNumberDiv1024(sheet, col++, row, gd[0].getVarValues("configurable_sys_mem"), 0, formatFloat(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				addNumberDiv1024(sheet, col++, row, gd[0].getVarValues("deconfig_sys_mem"), 0, formatFloat(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				addNumberDiv1024(sheet, col++, row, gd[0].getVarValues("sys_firmware_mem"), 0, formatFloat(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				addNumberDiv1024(sheet, col++, row, gd[0].getVarValues("curr_avail_sys_mem"), 0, formatFloat(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				addNumberDiv1024(sheet, col++, row, gd[0].getVarValues("pend_avail_sys_mem"), 0, formatFloat(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
			}
			
			addNumber(sheet, col++, row, managedSystem[i].getVarValues("sample_rate"),0, formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
			
			addLabel(sheet, col++, row, managedSystem[i].getVarValues("HMC1_name"),0, formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW|RIGHT));
			addLabel(sheet, col++, row, managedSystem[i].getVarValues("HMC2_name"),0, formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW|RIGHT));
			
			addLabel(sheet, col++, row, managedSystem[i].getVarValues("ipaddr"),0, formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW|RIGHT));
			addLabel(sheet, col++, row, managedSystem[i].getVarValues("ipaddr_secondary"),0, formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW|RIGHT));

			
			gd = managedSystem[i].getObjects(SYSPOWERLIC);
			if (gd!=null && gd[0]!=null) {	
				addLabel(sheet, col++, row, gd[0].getVarValues("ecnumber"),0, formatLabel(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW|RIGHT));
				addNumber(sheet, col++, row, gd[0].getVarValues("platform_ipl_level"),0, formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				addNumber(sheet, col++, row, gd[0].getVarValues("activated_level"),0, formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				if (gd[0].getVarValues("deferred_level")[0].equalsIgnoreCase("None"))
					addLabel(sheet, col++, row, gd[0].getVarValues("deferred_level"),0, formatLabel(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW|RIGHT));
				else
					addNumber(sheet, col++, row, gd[0].getVarValues("deferred_level"),0, formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
			}
			
			
			row++;
		}	
		
		sheet.setColumnView(0, nameSize+6);
		
		sheet.setColumnView(1, 9);
		sheet.setColumnView(2, 9);
		sheet.setColumnView(3, 9);
		sheet.setColumnView(4, 9);
		
		
		sheet.setColumnView(5, 6);
		sheet.setColumnView(6, 6);
		sheet.setColumnView(7, 6);
		
		sheet.setColumnView(8, 9);
		sheet.setColumnView(9, 9);
		
		sheet.setColumnView(10, 7);
		sheet.setColumnView(11, 7);
		sheet.setColumnView(12, 7);
		
		sheet.setColumnView(13, 7);
		
		sheet.setColumnView(14, 9);	
		sheet.setColumnView(15, 9);
		sheet.setColumnView(16, 9);
		sheet.setColumnView(17, 9);
		sheet.setColumnView(18, 9);
		sheet.setColumnView(19, 9);
		sheet.setColumnView(20, 9);
		
		sheet.setColumnView(21, 15);
		sheet.setColumnView(22, 15);	
		sheet.setColumnView(23, 15);
		sheet.setColumnView(24, 15);
		
		sheet.setColumnView(25, 9);
		sheet.setColumnView(26, 9);
		sheet.setColumnView(27, 9);
		sheet.setColumnView(28, 9);
	
	}
	
	
	
	private void createSystemsSheet(WritableSheet sheet) throws RowsExceededException, WriteException {
		int i;
		GenericData gd[];
		int col;
		int nameSize=0;
		String s;
			
		col = 0;
		
		/*
		 * Setup titles
		 */ 
		//sheet.mergeCells(0, row, 0, row+1);
		//addLabel(sheet,0,row,"Name",formatLabel(BOLD|VCENTRE|B_ALL_MED));
		
		sheet.mergeCells(0, 1, 1, 1);
		addLabel(sheet,0,1,"Status",formatLabel(BOLD|VCENTRE|B_ALL_MED));
		
		sheet.mergeCells(0, 2, 0, 3);
		addLabel(sheet,0,2,"Identification",formatLabel(BOLD|VCENTRE|B_ALL_MED));
		
		addLabel(sheet,1,2,"Type-Model",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,1,3,"Serial",formatLabel(BOLD|B_ALL_MED));
		
		sheet.mergeCells(0, 4, 0, 8);
		addLabel(sheet,0,4,"Cores",formatLabel(BOLD|VCENTRE|B_ALL_MED));
		
		addLabel(sheet,1,4,"Installed",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,1,5,"Active",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,1,6,"Deconfig",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,1,7,"Curr Avail",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,1,8,"Pend Avail",formatLabel(BOLD|B_ALL_MED));
		
		sheet.mergeCells(0, 9, 0, 14);
		addLabel(sheet,0,9,"Memory (MB)",formatLabel(BOLD|VCENTRE|B_ALL_MED));
		
		addLabel(sheet,1,9,"Installed",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,1,10,"Active",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,1,11,"Deconfig",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,1,12,"Firmware",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,1,13,"Curr Avail",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,1,14,"Pend Avail",formatLabel(BOLD|B_ALL_MED));
		
		sheet.mergeCells(0, 15, 1, 15);
		addLabel(sheet,0,15,"Perf Sample Rate",formatLabel(BOLD|B_ALL_MED));
		
		sheet.mergeCells(0, 16, 0, 17);
		addLabel(sheet,0,16,"Manager",formatLabel(BOLD|VCENTRE|B_ALL_MED));
		
		addLabel(sheet,1,16,"#1",formatLabel(BOLD|B_ALL_MED));
		addLabel(sheet,1,17,"#2",formatLabel(BOLD|B_ALL_MED));
		
		sheet.mergeCells(0, 18, 0, 19);
		addLabel(sheet,0,18,"Service Processor IP",formatLabel(BOLD|VCENTRE|B_ALL_MED|WRAP));
		
		addLabel(sheet,1,18,"Primary",formatLabel(BOLD|B_RIGHT_MED|B_TOP_MED|B_LEFT_MED));
		addLabel(sheet,1,19,"Secondary",formatLabel(BOLD|B_RIGHT_MED|B_BOTTOM_MED|B_LEFT_MED));
		
		sheet.mergeCells(0, 20, 0, 23);
		addLabel(sheet,0,20,"Code Levels",formatLabel(BOLD|VCENTRE|B_ALL_MED|WRAP));
		
		addLabel(sheet,1,20,"EC Number",formatLabel(BOLD|B_RIGHT_MED|B_TOP_MED|B_LEFT_MED));
		addLabel(sheet,1,21,"IPL Level",formatLabel(BOLD|B_RIGHT_MED|B_LEFT_MED));
		addLabel(sheet,1,22,"Activated Level",formatLabel(BOLD|B_RIGHT_MED|B_LEFT_MED));
		addLabel(sheet,1,23,"Deferred Level",formatLabel(BOLD|B_RIGHT_MED|B_BOTTOM_MED|B_LEFT_MED));

		
		col =2;
		
		
		/*
		 * Show values, each system on one line
		 */
		for (i=0; i<managedSystem.length; i++) {
			s = managedSystem[i].getVarValues("name")[0];
			if (s.length() > nameSize)
				nameSize = s.length();
			
			addLabel(sheet,col,0,s,formatLabel(BOLD|DIAG45));
			if (managerType == M_SDMC || managerType == M_FSM)
				addLabel(sheet,col,1,managedSystem[i].getVarValues("primary_state")[0],formatLabel(B_ALL_LOW));
			else
				addLabel(sheet,col,1,managedSystem[i].getVarValues("state")[0],formatLabel(B_ALL_LOW|RIGHT));
			addLabel(sheet,col,2,managedSystem[i].getVarValues("type_model")[0],formatLabel(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW|RIGHT));
			addLabel(sheet,col,3,managedSystem[i].getVarValues("serial_num")[0],formatLabel(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW|RIGHT));
			
			gd = managedSystem[i].getObjects(PROC);
			if (gd!=null && gd[0]!=null) {		
				addNumber(sheet,col,4,gd[0].getVarValues("installed_sys_proc_units"),0,formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				addNumber(sheet,col,5,gd[0].getVarValues("configurable_sys_proc_units"),0,formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				addNumber(sheet,col,6,gd[0].getVarValues("deconfig_sys_proc_units"),0,formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				addNumber(sheet,col,7, gd[0].getVarValues("curr_avail_sys_proc_units"),0, formatFloat(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				addNumber(sheet,col,8, gd[0].getVarValues("pend_avail_sys_proc_units"),0, formatFloat(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
			}
			
			gd = managedSystem[i].getObjects(MEM);
			if (gd!=null && gd[0]!=null) {	
				addNumber(sheet, col,  9, gd[0].getVarValues("installed_sys_mem"),0, formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				addNumber(sheet, col, 10, gd[0].getVarValues("configurable_sys_mem"),0, formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				addNumber(sheet, col, 11, gd[0].getVarValues("deconfig_sys_mem"),0, formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				addNumber(sheet, col, 12, gd[0].getVarValues("sys_firmware_mem"),0, formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				addNumber(sheet, col, 13, gd[0].getVarValues("curr_avail_sys_mem"),0, formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				addNumber(sheet, col, 14, gd[0].getVarValues("pend_avail_sys_mem"),0, formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
			}
			
			addNumber(sheet, col, 15, managedSystem[i].getVarValues("sample_rate"),0, formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
			
			addLabel(sheet, col, 16, managedSystem[i].getVarValues("HMC1_name"),0, formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW|RIGHT));
			addLabel(sheet, col, 17, managedSystem[i].getVarValues("HMC2_name"),0, formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW|RIGHT));
			
			addLabel(sheet, col, 18, managedSystem[i].getVarValues("ipaddr"),0, formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW|RIGHT));
			addLabel(sheet, col, 19, managedSystem[i].getVarValues("ipaddr_secondary"),0, formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW|RIGHT));
			//addLabel(sheet, col, 20, managedSystem[i].getVarValues("ipaddr2"),0, formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW|RIGHT));
			//addLabel(sheet, col, 21, managedSystem[i].getVarValues("ipaddr2_secondary"),0, formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW|RIGHT));
			
			gd = managedSystem[i].getObjects(SYSPOWERLIC);
			if (gd!=null && gd[0]!=null) {	
				addLabel(sheet, col, 20, gd[0].getVarValues("ecnumber"),0, formatLabel(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW|RIGHT));
				addNumber(sheet, col, 21, gd[0].getVarValues("platform_ipl_level"),0, formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				addNumber(sheet, col, 22, gd[0].getVarValues("activated_level"),0, formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
				if (gd[0].getVarValues("deferred_level")[0].equalsIgnoreCase("None"))
					addLabel(sheet, col, 23, gd[0].getVarValues("deferred_level"),0, formatLabel(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW|RIGHT));
				else
					addNumber(sheet, col, 23, gd[0].getVarValues("deferred_level"),0, formatInt(B_RIGHT_LOW|B_LEFT_LOW|B_BOTTOM_LOW));
			}
			
			
			col++;
		}
		
		// Set column size	
		for (i=0; i<=50; i++)
			sheet.setColumnView(i, 15);				
	}
	
	private void createGlobalSystemsSheet(WritableSheet sheet) throws RowsExceededException, WriteException {
		String names[];
		String s[];
		int i,j,k;
		GenericData gd[];

		
		// Create a cell format for Arial 10 point font
		WritableFont arial10font = new WritableFont(WritableFont.ARIAL, 10);
		WritableFont arial10boldfont = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD, true);
		WritableCellFormat arial10format = new WritableCellFormat (arial10font);
		WritableCellFormat arial10boldformat = new WritableCellFormat (arial10boldfont);
		
		WritableCellFormat diagonal = new WritableCellFormat (arial10boldfont);
		diagonal.setOrientation(Orientation.PLUS_45);
		
		WritableCellFormat centre = new WritableCellFormat (arial10boldfont);
		centre.setAlignment(Alignment.CENTRE);
		centre.setVerticalAlignment(VerticalAlignment.CENTRE);
		
		WritableCellFormat vertical = new WritableCellFormat (arial10boldfont);
		vertical.setOrientation(Orientation.PLUS_90);
		vertical.setAlignment(Alignment.CENTRE);
		vertical.setVerticalAlignment(VerticalAlignment.CENTRE);

		
		WritableCellFormat fraction = new WritableCellFormat (NumberFormats.FRACTION_TWO_DIGITS);
		WritableCellFormat integer = new WritableCellFormat (NumberFormats.INTEGER);
		
		
		
		Label label;
		Number number;
		String v;
		int row;
		
		
		// Set column size
		sheet.setColumnView(0, 5);
		sheet.setColumnView(1, 36);
		
		
		
		// Server names		
		for (i=0; i<managedSystem.length; i++) {
			label = new Label(2+i,0,managedSystem[i].getVarValues("name")[0],diagonal);
			sheet.addCell(label);
			sheet.setColumnView(2+i, 13);
		}
		sheet.getSettings().setVerticalFreeze(3);
		sheet.getSettings().setHorizontalFreeze(2);
		row = 1;
		
		// Type-Model ; Serial number ; State
		label = new Label(1,row,"Type-Model",arial10boldformat);
		sheet.addCell(label);
		label = new Label(1,row+1,"Serial",arial10boldformat);
		sheet.addCell(label);
		label = new Label(1,row+2,"State",arial10boldformat);
		sheet.addCell(label);
		for (i=0; i<managedSystem.length; i++) {
			label = new Label(2+i,row,managedSystem[i].getVarValues("type_model")[0],arial10format);
			sheet.addCell(label);
			label = new Label(2+i,row+1,managedSystem[i].getVarValues("serial_num")[0],arial10format);
			sheet.addCell(label);
			label = new Label(2+i,row+2,managedSystem[i].getVarValues("state")[0],arial10format);
			sheet.addCell(label);
		}
		row += 3;
	
		
		sheet.mergeCells(0, row-1, 0, row+4);
		label = new Label(0,row-1,"Summary",vertical);
		sheet.addCell(label);
		
		
		
		// CPU data		
		label = new Label(1,row,"Installed cores",arial10boldformat);
		sheet.addCell(label);
		label = new Label(1,row+1,"Deconfigured cores",arial10boldformat);
		sheet.addCell(label);
		for (i=0; i<managedSystem.length; i++) {
			gd = managedSystem[i].getObjects(PROC);
			if (gd==null || gd[0]==null)
				continue;
			number = new Number(2+i,row,Float.parseFloat(gd[0].getVarValues("installed_sys_proc_units")[0]),integer);
			sheet.addCell(number);
			number = new Number(2+i,row+1,Float.parseFloat(gd[0].getVarValues("deconfig_sys_proc_units")[0]),integer);
			sheet.addCell(number);
		}
		row += 2;
		
		// MEM data
		label = new Label(1,row,"Installed memory MB",arial10boldformat);
		sheet.addCell(label);
		label = new Label(1,row+1,"Deconfigured memory MB",arial10boldformat);
		sheet.addCell(label);
		label = new Label(1,row+2,"System Firmware memory MB",arial10boldformat);
		sheet.addCell(label);
		for (i=0; i<managedSystem.length; i++) {
			gd = managedSystem[i].getObjects(MEM);
			if (gd==null || gd[0]==null)
				continue;
			number = new Number(2+i,row,Float.parseFloat(gd[0].getVarValues("installed_sys_mem")[0]),integer);
			sheet.addCell(number);
			number = new Number(2+i,row+1,Float.parseFloat(gd[0].getVarValues("deconfig_sys_mem")[0]),integer);
			sheet.addCell(number);
			number = new Number(2+i,row+2,Float.parseFloat(gd[0].getVarValues("sys_firmware_mem")[0]),integer);
			sheet.addCell(number);
		}
		row += 3;
		
		row++;
		sheet.mergeCells(0, row, 1+managedSystem.length, row);
		label = new Label(0,row,"General properties",centre);
		sheet.addCell(label);
		row++;
		
		

		
		names = managedSystem[0].getVarNames();
		for (i=1; i<managedSystem.length; i++)
			names = mergeList(names, managedSystem[i].getVarNames());
		
		for (i=0; i<names.length; i++) {
			label = new Label(1,row+i,names[i],arial10boldformat);
			sheet.addCell(label);
		}		
		
		for (i=0; i<managedSystem.length; i++) {
			
			for (j=0; j<names.length; j++) {
				s = managedSystem[i].getVarValues(names[j]);
				if (s==null)
					continue;
				
				v=null;
				for (k=0; k<s.length; k++) {	
					if (v==null)
						v=s[k];
					else
						v=v+","+s[k];
				}
			
				label = new Label(2+i,row+j,v,arial10format);
				sheet.addCell(label);	
			}	
		}
		row += names.length;
		
		
		row++;
		sheet.mergeCells(0, row, 1+managedSystem.length, row);
		label = new Label(0,row,"Processor properties",centre);
		sheet.addCell(label);
		row++;
		

		
		names = managedSystem[0].getObjects(PROC)[0].getVarNames();
		for (i=1; i<managedSystem.length; i++) {
			gd = managedSystem[i].getObjects(PROC);
			if (gd!=null)
				names = mergeList(names, gd[0].getVarNames());
		}
		for (i=0; i<names.length; i++) {
			label = new Label(1,row+i,names[i],arial10boldformat);
			sheet.addCell(label);
		}
		for (i=0; i<managedSystem.length; i++) {
			
			for (j=0; j<names.length; j++) {
				gd = managedSystem[i].getObjects(PROC);
				if (gd==null)
					continue;
				s = gd[0].getVarValues(names[j]);
				if (s==null)
					continue;
				
				v=null;
				for (k=0; k<s.length; k++) {	
					if (v==null)
						v=s[k];
					else
						v=v+","+s[k];
				}
			
				label = new Label(2+i,row+j,v,arial10format);
				sheet.addCell(label);	
			}	
		}		
		row += names.length;
		
		
		row++;
		sheet.mergeCells(0, row, 1+managedSystem.length, row);
		label = new Label(0,row,"Memory properties",centre);
		sheet.addCell(label);
		row++;
		
		
		names = managedSystem[0].getObjects(MEM)[0].getVarNames();
		for (i=1; i<managedSystem.length; i++) {
			gd = managedSystem[i].getObjects(MEM);
			if (gd!=null)
				names = mergeList(names, gd[0].getVarNames());
		}
		for (i=0; i<names.length; i++) {
			label = new Label(1,row+i,names[i],arial10boldformat);
			sheet.addCell(label);
		}
		for (i=0; i<managedSystem.length; i++) {
			
			for (j=0; j<names.length; j++) {
				gd = managedSystem[i].getObjects(MEM);
				if (gd==null)
					continue;
				s = gd[0].getVarValues(names[j]);
				if (s==null)
					continue;
				
				v=null;
				for (k=0; k<s.length; k++) {	
					if (v==null)
						v=s[k];
					else
						v=v+","+s[k];
				}
			
				label = new Label(2+i,row+j,v,arial10format);
				sheet.addCell(label);	
			}	
		}
		
		
	}
	
	
	private void dumpSystemsSheet(WritableWorkbook w) throws RowsExceededException, WriteException {
		String names[];
		String s[];
		int i,j,k;
		GenericData gd[];
		
		WritableSheet sheet = workbook.createSheet("Systems", 0);
		
		// Create a cell format for Arial 10 point font
		WritableFont arial10font = new WritableFont(WritableFont.ARIAL, 10);
		WritableFont arial10boldfont = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD, true);
		WritableCellFormat arial10format = new WritableCellFormat (arial10font);
		WritableCellFormat arial10boldformat = new WritableCellFormat (arial10boldfont);
		
		WritableCellFormat diagonal = new WritableCellFormat (arial10boldfont);
		
		
		Label label;
		String v;
		int row;
		
		
		
		
		names = managedSystem[0].getVarNames();
		for (i=1; i<managedSystem.length; i++)
			names = mergeList(names, managedSystem[i].getVarNames());
		
		for (i=0; i<names.length; i++) {
			label = new Label(0,2+i,names[i],arial10boldformat);
			sheet.addCell(label);
		}
		
		diagonal.setOrientation(Orientation.PLUS_45);
		for (i=0; i<managedSystem.length; i++) {
			label = new Label(1+i,1,managedSystem[i].getVarValues("name")[0],diagonal);
			sheet.addCell(label);
		}
		
		
		
		for (i=0; i<managedSystem.length; i++) {
			
			for (j=0; j<names.length; j++) {
				s = managedSystem[i].getVarValues(names[j]);
				if (s==null)
					continue;
				
				v=null;
				for (k=0; k<s.length; k++) {	
					if (v==null)
						v=s[k];
					else
						v=v+","+s[k];
				}
			
				label = new Label(1+i,2+j,v,arial10format);
				sheet.addCell(label);	
			}	
		}	
		

		row = names.length+2;
		names = managedSystem[0].getObjects(PROC)[0].getVarNames();
		for (i=1; i<managedSystem.length; i++) {
			gd = managedSystem[i].getObjects(PROC);
			if (gd!=null)
				names = mergeList(names, gd[0].getVarNames());
		}
		for (i=0; i<names.length; i++) {
			label = new Label(0,row+i,names[i],arial10boldformat);
			sheet.addCell(label);
		}
		for (i=0; i<managedSystem.length; i++) {
			
			for (j=0; j<names.length; j++) {
				gd = managedSystem[i].getObjects(PROC);
				if (gd==null)
					continue;
				s = gd[0].getVarValues(names[j]);
				if (s==null)
					continue;
				
				v=null;
				for (k=0; k<s.length; k++) {	
					if (v==null)
						v=s[k];
					else
						v=v+","+s[k];
				}
			
				label = new Label(1+i,row+j,v,arial10format);
				sheet.addCell(label);	
			}	
		}
		
		row += names.length;
		names = managedSystem[0].getObjects(MEM)[0].getVarNames();
		for (i=1; i<managedSystem.length; i++) {
			gd = managedSystem[i].getObjects(MEM);
			if (gd!=null)
				names = mergeList(names, gd[0].getVarNames());
		}
		for (i=0; i<names.length; i++) {
			label = new Label(0,row+i,names[i],arial10boldformat);
			sheet.addCell(label);
		}
		for (i=0; i<managedSystem.length; i++) {
			
			for (j=0; j<names.length; j++) {
				gd = managedSystem[i].getObjects(MEM);
				if (gd==null)
					continue;
				s = gd[0].getVarValues(names[j]);
				if (s==null)
					continue;
				
				v=null;
				for (k=0; k<s.length; k++) {	
					if (v==null)
						v=s[k];
					else
						v=v+","+s[k];
				}
			
				label = new Label(1+i,row+j,v,arial10format);
				sheet.addCell(label);	
			}	
		}
		
		
	}
	
	
	
	
	private void identifyManagerType(String baseDir) {
		BufferedReader br;
		String line;		
		File f;
		
		f = new File(baseDir + lshmcV);	
		
		if (f.exists()) {
			try {
				br = new BufferedReader(new FileReader(baseDir + lshmcV),1024*1024);
				
				while ( (line = br.readLine()) != null ) {
					
					if (line.startsWith("HMC")) {
						managerType = M_HMC;
						return;
					}
					
					if (line.startsWith("SDMC")) {
						managerType = M_SDMC;
						return;
					}
					
					if (line.startsWith("FSM")) {
						managerType = M_FSM;
						return;
					}
				}
			} catch (IOException ioe) {	
				System.out.println("Loader.identifyManagerType: IOException");
				System.out.println(ioe);
			}	
		}
		
		f = new File(baseDir + ivmversion);	
		
		if (f.exists()) {
			try {
				br = new BufferedReader(new FileReader(baseDir + ivmversion),1024*1024);
				
				while ( (line = br.readLine()) != null ) {
					
					managerType = M_IVM;
					return;
				
				}
			} catch (IOException ioe) {	
				System.out.println("Loader.identifyManagerType: IOException");
				System.out.println(ioe);
			}	
		}
	}
	
	
	
	private void setScannerDate(String s) {
		int year, month, day;
		int s1,s2;
		
		s1 = s.indexOf('-');
		s2 = s.indexOf('-', s1+1);
		
		year = Integer.parseInt(s.substring(0, s1));
		month = Integer.parseInt(s.substring(s1+1, s2));
		day = Integer.parseInt(s.substring(s2+1));
				
		scannerDate = new GregorianCalendar(year, month-1, day);
	}
	
	
	
	
	
	private void loadScannerParams(String baseDir) {
		BufferedReader br;
		String names[]=null;
		String line;
		int j;
		DataParser dp;
		
		try {			
			br = new BufferedReader(new FileReader(baseDir + scannerInfo),1024*1024);
			
			names=null;
			while ( (line = br.readLine()) != null ) {
				
				dp = new DataParser(line);
				names = dp.getNames();
				scannerParams = new GenericData();
				
				for (j=0; j<names.length; j++) {
					scannerParams.add(names[j], dp.getStringValue(names[j]));					
				}		
			}	
			
			setScannerDate(scannerParams.getVarValues("HMCdate")[0]);
		} 
		catch (IOException ioe) {	
			System.out.println("Loader.loadSysConfigData.loadscannerParams: IOException");
			System.out.println(ioe);
		}
	}
	
	
	
	private void addManagedSystem(GenericData ms, String hmcName) {
		
		String nameS[] = new String[1];
		nameS[0] = hmcName;
		ms.add("HMC1_name",nameS);
			
		if (managedSystem == null) {
			managedSystem = new GenericData[1];
			managedSystem[0] = ms;
			return;
		}
 
		
		int curr,compare;
		String name = ms.getVarValues("name")[0];
		
		// Search position of new object
		curr=compare=0;
		while (curr<managedSystem.length && 
				(compare = managedSystem[curr].getVarValues("name")[0].compareTo(name)) <0 ) 
			curr++;
		
		// Name exists: just update HMC and SP 
		if (compare == 0) {
			managedSystem[curr].add("HMC2_name", nameS);
			
			String a[], b[];
			a = ms.getVarValues("ipaddr");
			b = managedSystem[curr].getVarValues("ipaddr");
			if (a[0].equals(b[0])) {
				// Two HMCs that point to same SP port: no update
			} else {
				// Second HMC uses second SP port
				managedSystem[curr].add("ipaddr2", ms.getVarValues("ipaddr"));
				managedSystem[curr].add("ipaddr2_secondary", ms.getVarValues("ipaddr_secondary"));
			}	
			return;
		}
		
		// Add new managed system
		GenericData newManagedSystem[] = new GenericData[managedSystem.length+1]; 
		int i;
		
		for (i=0; i<curr; i++)
			newManagedSystem[i] = managedSystem[i];
		newManagedSystem[curr] = ms;
		for (i=curr; i<managedSystem.length; i++)
			newManagedSystem[i+1] = managedSystem[i];
		managedSystem = newManagedSystem;			
	}
	
	
	
	private void loadSysConfigData(String hmcName, String baseDir) {
		BufferedReader br;
		String names[]=null;
		String line;
		int i,j;
		DataParser dp;
		
		// Temporary storage of managedSystem data
		//Vector<GenericData> msv = new Vector<GenericData>();
		GenericData ms = null;
		
		
		try {			
			br = new BufferedReader(new FileReader(baseDir + systemData),1024*1024);
			
			names=null;
			while ( (line = br.readLine()) != null ) {
				
				// Skip line if no data is returned
				if (line.startsWith("HSC") || line.startsWith("No results were found"))
					continue;
				
				dp = new DataParser(line);
				names = dp.getNames();
				ms = new GenericData();
				
				for (j=0; j<names.length; j++) {
					ms.add(names[j], dp.getStringValue(names[j]));					
				}
				//msv.add(ms);		
				addManagedSystem(ms, hmcName);
			}
			
			// Create managedSystem structure
			//managedSystem = new GenericData[msv.size()];
			//for (i=0; i<msv.size(); i++)
			//	managedSystem[i]=msv.elementAt(i);
			
			
			
			/*
			names=null;
			for (i=0; i<num; i++) {
				line = br.readLine();
				
				// Skip line if no data is returned
				if (line.startsWith("HSC"))
					continue;
				
				dp = new DataParser(line);
				names = dp.getNames();
				
				for (j=0; j<names.length; j++) {
					managedSystem[i].add(names[j], dp.getStringValue(names[j]));					
				}				
			}	
			*/
			
			
		} 
		catch (IOException ioe) {	
			System.out.println("Loader.loadSysConfigData.systemData: IOException");
			System.out.println(ioe);
		}
		
		
		// if no managed systems return 
		if (managedSystem == null)
			return;		
		
	
		
		// Add utilization data configuration
		Vector<String> system = new Vector<String>();
		Vector<String> sample = new Vector<String>();
		String model, serial;
		try {			
			br = new BufferedReader(new FileReader(baseDir + utilDataConfig),1024*1024);
			
			
			names=null;
			while ( (line = br.readLine()) != null ) {
				
				// Skip line if no data is returned
				if (line==null || line.startsWith("HSC") || !line.contains("type_model_serial_num"))
					continue;
				
				dp = new DataParser(line);
				system.add(dp.getStringValue("type_model_serial_num")[0]);
				sample.add(dp.getStringValue("sample_rate")[0]);							
			}
			

			/*
			names=null;
			for (i=0; i<num; i++) {
				line = br.readLine();
				
				// Skip line if no data is returned
				if (line==null || line.startsWith("HSC") || !line.contains("type_model_serial_num"))
					continue;
				
				dp = new DataParser(line);
				system.add(dp.getStringValue("type_model_serial_num")[0]);
				sample.add(dp.getStringValue("sample_rate")[0]);							
			}	
			*/
			
		} 
		catch (IOException ioe) {	
			System.out.println("Loader.loadSysConfigData: IOException.utilDataConfig");
			System.out.println(ioe);
		}
		
		
		for (i=0; i<managedSystem.length; i++) {
			model = managedSystem[i].getVarValues("type_model")[0];
			serial = managedSystem[i].getVarValues("serial_num")[0];
			
			for (j=0; j<system.size(); j++) {
				if (system.get(j).equals(model+"*"+serial)) {
					String result[] = new String[1];
					result[0]=sample.get(j);
					managedSystem[i].add("sample_rate", result);
					break;
				}
				
			}
			
		}		
	}
	
	
	/*
	 * Translate column ID to column name, zero based.
	 * col 0 ==> A, col 1 ==> B
	 * col 26 ==> AA
	 */
	private String getColName(int col) {
		int		C1 = col / 26;
		int		C2 = col % 26;
		String  result = "";
		
		if (C1>0)	
			result = String.valueOf((char)('A'+C1));
		result = result + String.valueOf((char)('A'+C2));
		
		return result;
	}
	
	
	
	private void sortManagedData() {
		int i,j;
		int id;
		String curr;
		
		GenericData gd;
		float[][] faa;
		boolean[] ba;
		
		for (i=0; i<managedSystem.length; i++) {
			curr = managedSystem[i].getVarValues("name")[0];
			id=-1;
			for (j=i+1; j<managedSystem.length; j++)
				//if (compareString(managedSystem[j].getVarValues("name")[0], curr)<0)
				//	id=j;
				if (managedSystem[j].getVarValues("name")[0].compareTo(curr)<0) {
					id=j;
					curr = managedSystem[j].getVarValues("name")[0];
				}
			if (id>=0) {
				gd = managedSystem[i];
				managedSystem[i]=managedSystem[id];
				managedSystem[id]=gd;
				
				faa=managedSystemData[i];
				managedSystemData[i]=managedSystemData[id];
				managedSystemData[id]=faa;
				
				ba = goodSystemData[i];
				goodSystemData[i]=goodSystemData[id];
				goodSystemData[id]=ba;		
				
				swapDataManagerObject(msCoreConfig,i,id);
				swapDataManagerObject(msCoreAvail,i,id);
				swapDataManagerObject(msCoreUsed,i,id);
				swapDataManagerObject(msMemConfig,i,id);
				swapDataManagerObject(msMemAvail,i,id);
			}
		}
	}
	
	
	private void swapDataManagerObject(DataManager[] array, int a, int b) {
		if (array==null || a<0 || a>array.length || b<0 || b>array.length)
			return;
		
		DataManager dm = array[a];
		array[a]=array[b];
		array[b] = dm;
	}
	
	
	private void swapLparStatusObject(NewLparStatus[] array, int a, int b) {
		if (array==null || a<0 || a>array.length || b<0 || b>array.length)
			return;
		
		NewLparStatus dm = array[a];
		array[a]=array[b];
		array[b] = dm;
	}
	
	
	private void sortLparData() {
		int i,j;
		int id;
		String curr;
		
		String s;
		float[][] faa;
		boolean[] ba;
		
		for (i=0; i<lparNames.length; i++) {
			curr = lparNames[i];
			id=-1;
			for (j=i+1; j<lparNames.length; j++)
				//if (compareString(lparNames[j],curr)<0) 
				//	id=j;
				if (lparNames[j].compareTo(curr)<0) {
					id=j;
					curr=lparNames[j];
				}
			if (id>=0) {
				s = lparNames[i];
				lparNames[i]=lparNames[id];
				lparNames[id]=s;
				
				/*
				faa=lparData[i];
				lparData[i]=lparData[id];
				lparData[id]=faa;
				
				ba = goodLparData[i];
				goodLparData[i]=goodLparData[id];
				goodLparData[id]=ba;	
				*/
				
				swapDataManagerObject(lparEnt,i,id);
				swapDataManagerObject(lparVP,i,id);
				swapDataManagerObject(lparPC,i,id);
				swapLparStatusObject(lparStatus,i,id);
			}
		}
	}
	
	
	private void sanitizeData() {
		int i,j,k;
		GenericData ms;
		GenericData[] gd, gd2;
		String[] s,s2,s3;
		String name;
		int n;
		
		StringChanger msChange = new StringChanger("System#");
		StringChanger snChange = new StringChanger("SN-");
		StringChanger ipChange = new StringChanger("IP#");
		StringChanger lparChange = new StringChanger("LPAR");
		StringChanger slotChange = new StringChanger("Drw#");
		StringChanger macChange = new StringChanger("MAC#");
		StringChanger vswitchChange = new StringChanger("VS#");
		StringChanger wwpnChange = new StringChanger("WWPN#");
		StringChanger vtdChange = new StringChanger("VTD#");
		
		StringChanger sc[] = new StringChanger[2];
		sc[0]=new StringChanger("VP#");
		sc[1]=msChange;
		StringChanger vpoolChange = new StringChanger("\n",sc);
				
		sanitizeString(scannerParams,"HMC",ipChange);
		sanitizeString(hmc,"serial",snChange);
		sanitizeString(hmc,"hostname",ipChange);
		sanitizeString(hmc,"gateway",ipChange);
		for (i=0; i<6; i++) {
			sanitizeString(hmc,"ipv4addr_eth"+(i-1),ipChange);
			sanitizeString(hmc,"ipv6addr_eth"+(i-1),ipChange);
		}
		
		for (i=0; i<managedSystem.length; i++) {
			ms = managedSystem[i];
			
			sanitizeString(ms,"name",msChange);
			sanitizeString(ms,"serial_num",snChange);
			sanitizeString(ms,"HMC1_name",ipChange);
			sanitizeString(ms,"HMC2_name",ipChange);
			sanitizeString(ms,"ipaddr",ipChange);
			sanitizeString(ms,"ipaddr_secondary",ipChange);			
			
			
			gd = ms.getObjects(CONFIG_LPAR);
			for (j=0; gd!=null && j<gd.length; j++) {
				sanitizeString(gd[j],"name",lparChange);
				sanitizeString(gd[j],"rmc_ipaddr",ipChange);
			}
			
			gd = ms.getObjects(PROC_LPAR);
			for (j=0; gd!=null && j<gd.length; j++) {
				sanitizeString(gd[j],"lpar_name",lparChange);
			}
			
			gd = ms.getObjects(MEM_LPAR);
			for (j=0; gd!=null && j<gd.length; j++) {
				sanitizeString(gd[j],"lpar_name",lparChange);
				sanitizeString(gd[j],"primary_paging_vios_name",lparChange);
				sanitizeString(gd[j],"secondary_paging_vios_name",lparChange);
				sanitizeString(gd[j],"curr_paging_vios_name",lparChange);
			}
			
			gd = ms.getObjects(SLOT);
			for (j=0; gd!=null && j<gd.length; j++) {
				sanitizeString(gd[j],"lpar_name",lparChange);
				
				s = gd[j].getVarValues("drc_name");
				n = s[0].indexOf('-');
				if (n>=0) {
					s[0] = slotChange.translate(s[0].substring(0, n)) + s[0].substring(n);
					gd[j].add("drc_name", s);
				}
			}
			
			gd = ms.getObjects(VETH);
			for (j=0; gd!=null && j<gd.length; j++) {
				sanitizeString(gd[j],"lpar_name",lparChange);
				sanitizeString(gd[j],"mac_addr",macChange);
				sanitizeString(gd[j],"vswitch",vswitchChange);
			}
			
			gd = ms.getObjects(VSWITCH);
			for (j=0; gd!=null && j<gd.length; j++) {
				if (gd[j]==null)
					continue;
				s = gd[j].getVarValues("vswitch");
				if (s[0].endsWith("(Default)")) {
					s[0]=s[0].substring(0, s[0].length()-9);
					s[0] = vswitchChange.translate(s[0]);
					s[0] = s[0]+"(Default)";
					gd[j].add("vswitch", s);
					continue;
				}
				sanitizeString(gd[j],"vswitch",vswitchChange);
			}
			
			gd = ms.getObjects(VSCSI);
			for (j=0; gd!=null && j<gd.length; j++) {
				sanitizeString(gd[j],"lpar_name",lparChange);
				sanitizeString(gd[j],"remote_lpar_name",lparChange);
			}
			
			gd = ms.getObjects(VSCSIMAP);
			for (j=0; gd!=null && j<gd.length; j++) {
				sanitizeString(gd[j],"VIOS",lparChange);
				sanitizeString(gd[j],"remote_lpar_name",lparChange);
				
				s = gd[j].getVarValues("physloc");
				if (s==null)
					continue;
				for (k=0; k<s.length; k++) {
					if (s[k]==null)
						continue;
					n = s[k].indexOf('-');
					if (n>=0) {
						s[k] = slotChange.translate(s[k].substring(0, n)) + s[k].substring(n);
					}
				}
				gd[j].add("physloc", s);
				
				s = gd[j].getVarValues("VTD");
				if (s==null)
					continue;
				for (k=0; k<s.length; k++) {
					if (s[k]==null)
						continue;
					s[k] = vtdChange.translate(s[k]);
				}
				gd[j].add("VTD", s);
				
				
			}
			
			gd = ms.getObjects(HDISK);
			for (j=0; gd!=null && j<gd.length; j++) {
				sanitizeString(gd[j],"VIOS",lparChange);
			}
			
			gd = ms.getObjects(VFC);
			for (j=0; gd!=null && j<gd.length; j++) {
				sanitizeString(gd[j],"lpar_name",lparChange);
				sanitizeString(gd[j],"remote_lpar_name",lparChange);
				
				s = gd[j].getVarValues("wwpns");
				if (s!=null) {
					s[0] = wwpnChange.translate(s[0]);
					s[1] = wwpnChange.translate(s[0]);
					gd[j].add("wwpns", s);
				}
			}
			
			gd = ms.getObjects(VFCMAP);
			for (j=0; gd!=null && j<gd.length; j++) {
				s = gd[j].getVarNames();
				for (k=0; s!=null && k<s.length; k++) {
					name = lparChange.translate(s[k].substring(0, s[k].indexOf('@'))) + s[k].substring(s[k].indexOf('@'));
					s2 = gd[j].getVarValues(s[k]);
					s3 = new String[1];
					n = s2[0].indexOf('-');
					if (n>=0) {
						s3[0] = slotChange.translate(s2[0].substring(0, n)) + s2[0].substring(n);
						gd[j].add(name, s3);
					}
				}
				
			}
		}
		
		if (lparNames!=null)
			for (i=0; i<lparNames.length; i++) {
				lparNames[i] = lparChange.translate(lparNames[i]);
			}
		
		if (procPoolName!=null)
			for (i=0; i<procPoolName.length; i++) {
				procPoolName[i] = vpoolChange.translate(procPoolName[i]);
			}
		
		if (lparStatus!=null) {
			for (i=0; i<lparStatus.length; i++) {
				lparStatus[i].sanitize(msChange, vpoolChange);
			}
		}
		
		for (i=0; i<managedSystem.length; i++) {
			ms = managedSystem[i];
			gd = ms.getObjects(PROC_POOL);
			for (j=0; gd!=null && j<gd.length; j++) {
				sanitizeString(gd[j],"name",vpoolChange);
			}
		}
	}
	
	
	private void sanitizeString(GenericData gd, String name, StringChanger sc) {
		String s[] = gd.getVarValues(name);
		if (s==null)
			return;
		s[0] = sc.translate(s[0]);
		gd.add(name, s);
	}
	
	
	
	/*
	 * Compare string: 
	 * 		<0		a precedes b
	 * 		0		a=b
	 * 		>0		b precedes a
	 */
	private int compareString(String a, String b) {
		int i;
		
		for (i=0; i<a.length() && i<b.length(); i++) {
			if (a.charAt(i)<b.charAt(i))
				return -1;
			if (a.charAt(i)>b.charAt(i))
				return 1;
		}
		
		return a.length()-b.length();
	}
	
	
	
	/*
	 * Fix the case where number is "null". Translate with 0
	 */
	private int textToInt (String s) {
		if (s.equals("null"))
			return 0;
		return Integer.parseInt(s);
	}
	
	
	public String getHMCdate() {
		String s[];
		String result;
		
		if (scannerParams == null)
			return "";
		
		s = scannerParams.getVarValues("HMCdate");
		if (s!=null && s.length>=1)
			result = s[0].replace("-","");
		else
			return "";
		
		s = scannerParams.getVarValues("HMCtime");
		if (s!=null && s.length>=1) {
			result = result + "_" + s[0].replace(":", "");
		} else
			return "";
		
		return result;
	}
	
	
	
	
	private void createIndexHtml(String dirName) {
		PrintWriter index;	
		
		try {
			index = new PrintWriter(
						new FileOutputStream(dirName + File.separatorChar + index_html));
		}
		catch (IOException e) { 
			System.out.println("\nError creating file "+dirName + File.separatorChar + index_html);
			System.out.println("Skipping HTML creation");
			return; 
		}
		
		// Create iframed index
		index.println("<HTML>\n<HEAD>\n<TITLE>hmcScanner HTML</TITLE>");	
		index.println("<style type=\"text/css\">");	
		index.println("<!--\niframe {\nfloat:left;\nmargin:0px;\n}\n-->");
		index.println("</style>\n</head>");
		index.println("<BODY BCOLOR=#e6e6e6>");
		index.println("<iframe width=\"15%\" height=\"100%\" src=\""+menu_html+"\" frameborder=\"0\" name=\"menu\"></iframe>");
		index.println("<iframe width=\"85%\" height=\"100%\" src=\""+header_html+"\" frameborder=\"0\" name=\"data\"</iframe>");
		index.println("</BODY>");
		index.println("</HTML>");
		index.close();		
	}
	
	private void createSystemPerfIndexHtml(String dirName) {
		PrintWriter index;	
		
		try {
			index = new PrintWriter(
						new FileOutputStream(dirName + File.separatorChar + sysperfindex_html));
		}
		catch (IOException e) { 
			System.out.println("\nError creating file "+dirName + File.separatorChar + sysperfindex_html);
			System.out.println("Skipping HTML creation");
			return; 
		}
		
		// Create iframed index
		index.println("<HTML>\n<HEAD>\n");	
		index.println("<style type=\"text/css\">");	
		index.println("<!--\niframe {\nfloat:left;\nmargin:0px;\n}\n-->");
		index.println("</style>\n</head>");
		index.println("<BODY BCOLOR=#e6e6e6>");
		index.println("<iframe width=\"15%\" height=\"100%\" src=\""+sysperfmenu_html+"\" frameborder=\"0\" name=\"sysperfmenu\"></iframe>");
		index.println("<iframe width=\"85%\" height=\"100%\" frameborder=\"0\" name=\"sysperfdata\"</iframe>");
		index.println("</BODY>");
		index.println("</HTML>");
		index.close();		
	}
	
	private void createPoolPerfIndexHtml(String dirName) {
		PrintWriter index;	
		
		try {
			index = new PrintWriter(
						new FileOutputStream(dirName + File.separatorChar + poolperfindex_html));
		}
		catch (IOException e) { 
			System.out.println("\nError creating file "+dirName + File.separatorChar + poolperfindex_html);
			System.out.println("Skipping HTML creation");
			return; 
		}
		
		// Create iframed index
		index.println("<HTML>\n<HEAD>\n");	
		index.println("<style type=\"text/css\">");	
		index.println("<!--\niframe {\nfloat:left;\nmargin:0px;\n}\n-->");
		index.println("</style>\n</head>");
		index.println("<BODY BCOLOR=#e6e6e6>");
		index.println("<iframe width=\"15%\" height=\"100%\" src=\""+poolperfmenu_html+"\" frameborder=\"0\" name=\"sysperfmenu\"></iframe>");
		index.println("<iframe width=\"85%\" height=\"100%\" frameborder=\"0\" name=\"poolperfdata\"</iframe>");
		index.println("</BODY>");
		index.println("</HTML>");
		index.close();		
	}
	
	
	private void createLPARPerfIndexHtml(String dirName) {
		PrintWriter index;	
		
		try {
			index = new PrintWriter(
						new FileOutputStream(dirName + File.separatorChar + lparperfindex_html));
		}
		catch (IOException e) { 
			System.out.println("\nError creating file "+dirName + File.separatorChar + lparperfindex_html);
			System.out.println("Skipping HTML creation");
			return; 
		}
		
		// Create iframed index
		index.println("<HTML>\n<HEAD>\n");	
		index.println("<style type=\"text/css\">");	
		index.println("<!--\niframe {\nfloat:left;\nmargin:0px;\n}\n-->");
		index.println("</style>\n</head>");
		index.println("<BODY BCOLOR=#e6e6e6>");
		index.println("<iframe width=\"15%\" height=\"100%\" src=\""+lparperfmenu_html+"\" frameborder=\"0\" name=\"sysperfmenu\"></iframe>");
		index.println("<iframe width=\"85%\" height=\"100%\" frameborder=\"0\" name=\"lparperfdata\"</iframe>");
		index.println("</BODY>");
		index.println("</HTML>");
		index.close();		
	}
	
	private void addButton(String button, String html) {
		if (buttonName==null) {
			buttonName = new Vector<String>();
			htmlName = new Vector<String>();
		}
		buttonName.add(button);
		htmlName.add(html);
	}
	
	private void addButtonSystem(String button, String html) {
		if (sysButtonName==null) {
			sysButtonName = new Vector<String>();
			sysHtmlName = new Vector<String>();
		}
		sysButtonName.add(button);
		sysHtmlName.add(html);
	}
	
	private void addButtonLPAR(String button, String html) {
		if (lparButtonName==null) {
			lparButtonName = new Vector<String>();
			lparHtmlName = new Vector<String>();
		}
		lparButtonName.add(button);
		lparHtmlName.add(html);
	}
	
	private void addButtonPool(String button, String html) {
		if (poolButtonName==null) {
			poolButtonName = new Vector<String>();
			poolHtmlName = new Vector<String>();
		}
		poolButtonName.add(button);
		poolHtmlName.add(html);
	}
	
	
	private void createSystemHtmlStructure(String dirName) {
		File f;
		int i;
		PrintWriter target;
		String name;
		
		// Create structure
		createSystemPerfIndexHtml(dirName);
		
		// Create graphic and related html + button structure
		for (i=0; i<managedSystem.length; i++) {
			name = managedSystem[i].getVarValues("name")[0];
			
			// Create target page
			try {
				target = new PrintWriter(
							new FileOutputStream(
								new File(dirName + File.separatorChar + name + ".html")));
			}
			catch (IOException e) { 
				System.out.println("\nError creating file "+dirName + File.separatorChar + name + ".html");
				System.out.println("Skipping HTML creation");
				return; 
			}
			
			target.println("<HTML><HEAD></HEAD>");
			target.println("<BODY BGCOLOR=#e6e6e6 TEXT=#000000 LINK=#0000FF VLINK=#0000FF ALINK=#FF0000><H1>"+name+"</H1>");
			target.println("<IMG SRC="+name+"_daily.png><BR>");
			target.println("<BR><BR>");
			target.println("<IMG SRC="+name+"_hourly.png><BR>");
			target.println("</BODY></HTML");
			target.close();
					
			f = new File(dirName + File.separatorChar + name+"_daily.png");
			createSysPoolDailyImage(i,f);
			
			f = new File(dirName + File.separatorChar + name+"_hourly.png");
			createSysPoolHourlyImage(i,f);
			
			addButtonSystem(name,name + ".html");
		}
		
		// Create menu with buttons
		createSystemMenuHtml(dirName);
	}
	
	
	private void createLPARHtmlStructure(String dirName) {
		File f;
		int i;
		PrintWriter target;
		String name;
		
		// Create structure
		createLPARPerfIndexHtml(dirName);
		
		// Create graphic and related html + button structure
		for (i=0; i<lparNames.length; i++) {
			name = lparNames[i];
			
			// Create target page
			try {
				target = new PrintWriter(
							new FileOutputStream(
								new File(dirName + File.separatorChar + name + ".html")));
			}
			catch (IOException e) { 
				System.out.println("\nError creating file "+dirName + File.separatorChar + name + ".html");
				System.out.println("Skipping HTML creation");
				return; 
			}
			
			target.println("<HTML><HEAD></HEAD>");
			target.println("<BODY BGCOLOR=#e6e6e6 TEXT=#000000 LINK=#0000FF VLINK=#0000FF ALINK=#FF0000><H1>"+name+"</H1>");
			target.println("<IMG SRC="+name+"_daily.png><BR>");
			target.println("<BR><BR>");
			target.println("<IMG SRC="+name+"_hourly.png><BR>");
			target.println("</BODY></HTML");
			target.close();
					
			f = new File(dirName + File.separatorChar + name+"_daily.png");
			createLparDailyImage(i,f);
			
			f = new File(dirName + File.separatorChar + name+"_hourly.png");
			createLparHourlyImage(i,f);
			
			addButtonLPAR(name,name + ".html");
		}
		
		// Create menu with buttons
		createLPARMenuHtml(dirName);
	}
	
	
	private void createPoolHtmlStructure(String dirName) {
		File f;
		int i;
		PrintWriter target;
		String name;
		
		// Create structure
		createPoolPerfIndexHtml(dirName);
		
		// Create graphic and related html + button structure
		for (i=0; procPoolName!=null && i<procPoolName.length; i++) {
			name = procPoolName[i].replace('\n', '@');
			
			// Create target page
			try {
				target = new PrintWriter(
							new FileOutputStream(
								new File(dirName + File.separatorChar + name + ".html")));
			}
			catch (IOException e) { 
				System.out.println("\nError creating file "+dirName + File.separatorChar + name + ".html");
				System.out.println("Skipping HTML creation");
				return; 
			}
			
			target.println("<HTML><HEAD></HEAD>");
			target.println("<BODY BGCOLOR=#e6e6e6 TEXT=#000000 LINK=#0000FF VLINK=#0000FF ALINK=#FF0000><H1>"+name+"</H1>");
			target.println("<IMG SRC="+name+"_daily.png><BR>");
			target.println("<BR><BR>");
			target.println("<IMG SRC="+name+"_hourly.png><BR>");
			target.println("</BODY></HTML");
			target.close();
					
			f = new File(dirName + File.separatorChar + name+"_daily.png");
			createVirtPoolDailyImage(i,f);
			
			f = new File(dirName + File.separatorChar + name+"_hourly.png");
			createVirtPoolHourlyImage(i,f);
			
			addButtonPool(name,name + ".html");
		}
		
		// Create menu with buttons
		createPoolMenuHtml(dirName);
	}
	
	private void createPoolMenuHtml(String dirName) {
		PrintWriter index;
		int i;
		
		try {
			index = new PrintWriter(
						new FileOutputStream(dirName + File.separatorChar + poolperfmenu_html));
		}
		catch (IOException e) { 
			System.out.println("\nError creating file "+dirName + File.separatorChar + poolperfmenu_html);
			System.out.println("Skipping HTML creation");
			return; 
		}
		
		// Create iframed index
		index.println("<HTML>\n<HEAD>");
		index.println("<STYLE>\n\tbutton{width:100%;font-family:calibri;font-size:12px;}\n</STYLE>");
		index.println("</HEAD>");
		index.println("<BODY BGCOLOR=#e6e6e6 TEXT=#000000 LINK=#0000FF VLINK=#0000FF ALINK=#FF0000>");
		
		// <input type="button" value="mioBottone" onclick="parent.graphframe.location.href='hmc.html'">
		
		for (i=0; poolButtonName!=null && i<poolButtonName.size(); i++) {
			index.println("<button type=\"button\"" +
					" onclick=\"parent.poolperfdata.location.href='"+poolHtmlName.elementAt(i)+"'\">");
			index.println(poolButtonName.elementAt(i)+"</button>");
			index.println("<BR>");
		}
		index.println("</BODY>");
		index.println("</HTML>");
		index.close();		
	}
	
	private void createSystemMenuHtml(String dirName) {
		PrintWriter index;
		int i;
		
		try {
			index = new PrintWriter(
						new FileOutputStream(dirName + File.separatorChar + sysperfmenu_html));
		}
		catch (IOException e) { 
			System.out.println("\nError creating file "+dirName + File.separatorChar + sysperfmenu_html);
			System.out.println("Skipping HTML creation");
			return; 
		}
		
		// Create iframed index
		index.println("<HTML>\n<HEAD>");
		index.println("<STYLE>\n\tbutton{width:100%;font-family:calibri;font-size:12px;}\n</STYLE>");
		index.println("</HEAD>");
		index.println("<BODY BGCOLOR=#e6e6e6 TEXT=#000000 LINK=#0000FF VLINK=#0000FF ALINK=#FF0000>");
		
		// <input type="button" value="mioBottone" onclick="parent.graphframe.location.href='hmc.html'">
		
		for (i=0; i<sysButtonName.size(); i++) {
			index.println("<button type=\"button\"" +
					" onclick=\"parent.sysperfdata.location.href='"+sysHtmlName.elementAt(i)+"'\">");
			index.println(sysButtonName.elementAt(i)+"</button>");
			index.println("<BR>");
		}
		index.println("</BODY>");
		index.println("</HTML>");
		index.close();		
	}
	
	private void createLPARMenuHtml(String dirName) {
		PrintWriter index;
		int i;
		
		try {
			index = new PrintWriter(
						new FileOutputStream(dirName + File.separatorChar + lparperfmenu_html));
		}
		catch (IOException e) { 
			System.out.println("\nError creating file "+dirName + File.separatorChar + lparperfmenu_html);
			System.out.println("Skipping HTML creation");
			return; 
		}
		
		// Create iframed index
		index.println("<HTML>\n<HEAD>");
		index.println("<STYLE>\n\tbutton{width:100%;font-family:calibri;font-size:12px;}\n</STYLE>");
		index.println("</HEAD>");
		index.println("<BODY BGCOLOR=#e6e6e6 TEXT=#000000 LINK=#0000FF VLINK=#0000FF ALINK=#FF0000>");
		
		// <input type="button" value="mioBottone" onclick="parent.graphframe.location.href='hmc.html'">
		
		for (i=0; i<lparButtonName.size(); i++) {
			index.println("<button type=\"button\"" +
					" onclick=\"parent.lparperfdata.location.href='"+lparHtmlName.elementAt(i)+"'\">");
			index.println(lparButtonName.elementAt(i)+"</button>");
			index.println("<BR>");
		}
		index.println("</BODY>");
		index.println("</HTML>");
		index.close();		
	}
	
	private void createMenuHtml(String dirName) {
		PrintWriter index;
		int i;
		
		try {
			index = new PrintWriter(
						new FileOutputStream(dirName + File.separatorChar + menu_html));
		}
		catch (IOException e) { 
			System.out.println("\nError creating file "+dirName + File.separatorChar + menu_html);
			System.out.println("Skipping HTML creation");
			return; 
		}

		
		// Create iframed index
		index.println("<HTML>\n<HEAD>");
		index.println("<STYLE>\n\tbutton{width:100%;font-family:calibri;font-size:12px;}\n</STYLE>");
		index.println("</HEAD>");
		index.println("<BODY BGCOLOR=#e6e6e6 TEXT=#000000 LINK=#0000FF VLINK=#0000FF ALINK=#FF0000>");
		
		// <input type="button" value="mioBottone" onclick="parent.graphframe.location.href='hmc.html'">
		
		for (i=0; i<buttonName.size(); i++) {
			index.println("<button type=\"button\"" +
					" onclick=\"parent.data.location.href='"+htmlName.elementAt(i)+"'\">");
			index.println(buttonName.elementAt(i)+"</button>");
			index.println("<BR>");
		}
		index.println("</BODY>");
		index.println("</HTML>");
		index.close();		
	}
	
	private void createHTML(String dirName) {
		File f;
		int i;
		PrintWriter index;
		PrintWriter target;
		
		
		// Create directory
	    f = new File(dirName);
		if (!f.isDirectory() && !f.mkdir()) {
			System.out.println("Error: can not create directory "+dirName);
			System.out.println("Skipping HTML creation");
			return;
		}
		
		System.out.print("Starting HTML file creation: ");
		
		
		
		try {
			index = new PrintWriter(
						new FileOutputStream(dirName + File.separatorChar + "index.html"));
		}
		catch (IOException e) { 
			System.out.println("\nError creating file "+dirName + File.separatorChar + "index.html");
			System.out.println("Skipping HTML creation");
			return; 
		}
		
		// Create framed index
		index.println("<HTML><HEAD><TITLE>hmcScanner graphs</TITLE></HEAD>");	
		index.println("<frameset framespacing=\"0\" border=\"0\" rows=\"105,*\" frameborder=\"1\">");
		index.println("  <frame name=\"topframe\" scrolling=\"no\" noresize src=\"top.html\" marginwidth=\"0\" marginheight=\"0\">");
		index.println("  <frameset cols=\"215,*\">");
		index.println("    <frame name=\"menuframe\" target=\"graphframe\" src=\"menu.html\" marginwidth=\"0\" marginheight=\"0\" scrolling=\"auto\">");
		index.println("    <frame name=\"graphframe\" src=\"welcome.html\" marginwidth=\"12\" marginheight=\"16\" scrolling=\"auto\">");
		index.println("  </frameset>");
		index.println("</frameset>");
		index.println("</HTML>");
		index.close();		
		
		// Create welcome page
		try {
			index = new PrintWriter(
						new FileOutputStream(
							new File(dirName + File.separatorChar + "welcome.html")));
		}
		catch (IOException e) { 
			System.out.println("\nError creating file "+dirName + File.separatorChar + "welcome.html");
			System.out.println("Skipping HTML creation");
			return; 
		}
		
		index.println("<HTML><HEAD><TITLE>Menu</TITLE></HEAD>");
		index.println("<BODY BGCOLOR=#e6e6e6 TEXT=#000000 LINK=#0000FF VLINK=#0000FF ALINK=#FF0000>");
		index.println("<H1>Welcome to hmcScanner generated reports.</H1>");
		index.println("Please select on the left frame the item you want to see.");
		index.println("</BODY></HTML>");
		index.close();	
		
		// Create top
		try {
			index = new PrintWriter(
						new FileOutputStream(
							new File(dirName + File.separatorChar + "top.html")));
		}
		catch (IOException e) { 
			System.out.println("\nError creating file "+dirName + File.separatorChar + "top.html");
			System.out.println("Skipping HTML creation");
			return; 
		}


	
		index.println("<HTML>" +
						"<HEAD>" +
							"<TITLE>Top</TITLE>" +
						"</HEAD>");
		index.println("<BODY BGCOLOR=#e6e6e6 TEXT=#000000 LINK=#0000FF VLINK=#0000FF ALINK=#FF0000>" +
						"<CENTER>" +
							"<H1>");
		index.println("hmcScanner Top ");							

		index.println("\n</H1></CENTER>\n");
		index.println("</HTML>");
		index.close();
		
		try {
			index = new PrintWriter(
						new FileOutputStream(
							new File(dirName + File.separatorChar + "menu.html")));
		}
		catch (IOException e) { 
			System.out.println("\nError creating file "+dirName + File.separatorChar + "menu.html");
			System.out.println("Skipping HTML creation");
			return; 
		}

		
		// Create menu
		index.println("<HTML><HEAD><TITLE>Menu</TITLE></HEAD>");
		index.println("<BODY BGCOLOR=#e6e6e6 TEXT=#000000 LINK=#0000FF VLINK=#0000FF ALINK=#FF0000>");
		index.println("<BR>");
		
		for (i=0; i<lparNames.length; i++) {
			index.println("<LI><A HREF=\""+i+".html"+"\" target=\"graphframe\">"+lparNames[i]+"</A></LI>");
			
			// Create target page
			try {
				target = new PrintWriter(
							new FileOutputStream(
								new File(dirName + File.separatorChar + i + ".html")));
			}
			catch (IOException e) { 
				System.out.println("\nError creating file "+dirName + File.separatorChar + i + ".html");
				System.out.println("Skipping HTML creation");
				return; 
			}
			
			target.println("<HTML><HEAD><TITLE>"+lparNames[i]+"</TITLE></HEAD>");
			target.println("<BODY BGCOLOR=#e6e6e6 TEXT=#000000 LINK=#0000FF VLINK=#0000FF ALINK=#FF0000><H1>"+lparNames[i]+"</H1>");
			target.println("<IMG SRC="+i+".D.png><BR>");
			target.println("<BR><BR>");
			target.println("<IMG SRC="+i+".H.png><BR>");
			target.println("</BODY></HTML");
			target.close();
					
			f = new File(dirName + File.separatorChar + i + ".D.png");
			createLparDailyImage(i,f);
			f = new File(dirName + File.separatorChar + i + ".H.png");
			createLparHourlyImage(i,f);
		}
		
		
		for (i=0; i<managedSystem.length; i++) {
			index.println("<LI><A HREF=\""+i+".P.html"+"\" target=\"graphframe\">"+managedSystem[i].getVarValues("name")[0]+"</A></LI>");
			
			// Create target page
			try {
				target = new PrintWriter(
							new FileOutputStream(
								new File(dirName + File.separatorChar + i + ".P.html")));
			}
			catch (IOException e) { 
				System.out.println("\nError creating file "+dirName + File.separatorChar + i + ".P.html");
				System.out.println("Skipping HTML creation");
				return; 
			}
			
			target.println("<HTML><HEAD><TITLE>"+managedSystem[i].getVarValues("name")[0]+"</TITLE></HEAD>");
			target.println("<BODY BGCOLOR=#e6e6e6 TEXT=#000000 LINK=#0000FF VLINK=#0000FF ALINK=#FF0000><H1>"+managedSystem[i].getVarValues("name")[0]+"</H1>");
			target.println("<IMG SRC="+i+".P.D.png><BR>");
			target.println("<BR><BR>");
			target.println("<IMG SRC="+i+".P.H.png><BR>");
			target.println("</BODY></HTML");
			target.close();
					
			f = new File(dirName + File.separatorChar + i + ".P.D.png");
			createSysPoolDailyImage(i,f);
			f = new File(dirName + File.separatorChar + i + ".P.H.png");
			createSysPoolHourlyImage(i,f);
		}
		
		
		for (i=0; procPoolName!=null && i<procPoolName.length; i++) {
			index.println("<LI><A HREF=\""+i+".VP.html"+"\" target=\"graphframe\">"+procPoolName[i]+"</A></LI>");
			
			// Create target page
			try {
				target = new PrintWriter(
							new FileOutputStream(
								new File(dirName + File.separatorChar + i + ".VP.html")));
			}
			catch (IOException e) {  
				System.out.println("\nError creating file "+dirName + File.separatorChar + i + ".VP.html");
				System.out.println("Skipping HTML creation");
				return; 
			}
			
			target.println("<HTML><HEAD><TITLE>"+procPoolName[i]+"</TITLE></HEAD>");
			target.println("<BODY BGCOLOR=#e6e6e6 TEXT=#000000 LINK=#0000FF VLINK=#0000FF ALINK=#FF0000><H1>"+procPoolName[i]+"</H1>");
			target.println("<IMG SRC="+i+".VP.D.png><BR>");
			target.println("<BR><BR>");
			target.println("<IMG SRC="+i+".VP.H.png><BR>");
			target.println("</BODY></HTML");
			target.close();
					
			f = new File(dirName + File.separatorChar + i + ".VP.D.png");
			createVirtPoolDailyImage(i,f);
			f = new File(dirName + File.separatorChar + i + ".VP.H.png");
			createVirtPoolHourlyImage(i,f);
		}
		
		
		index.println("</BODY></HTML>");
		index.close();
		
		System.out.println(" DONE. See " + dirName);
		
	}
	
	
	
	
	private void createLparDailyFiles(String dirName) {
		int i;
		File f, dir;
		
		// Create directory
	    dir = new File(dirName);
		if (!dir.isDirectory() && !dir.mkdir()) {
			System.out.println("Error: can not create directory "+dirName);
			System.exit(1);
		}
		
		System.out.print("Starting HTML file creation: ");
		
		for (i=0; i<lparNames.length; i++) {
			f = new File(dirName + File.separatorChar + i + ".png");
			createLparDailyImage(i,f);
		}
		
		System.out.println(" Done!");
		
	}
	
	
	
	private void createLparHourlyImage(int id, File file) {
		float pc[] = new float[24*60];
		float ent[] = new float[24*60];
		float vp[] = new float[24*60];
		float max[] = new float[24*60];
		
		final int					STEP = 5*24;
		String label[] = new String[24*60/STEP];
		int i,j;
		LparGraph lg=null;
		
		boolean cap;
		String ms;
		String pool;
		float free;
		int poolID;
		GregorianCalendar gc;
		
		for (i=0; i<pc.length; i++) {
			pc[i] = lparPC[id].getHourData(i);
			ent[i] = lparEnt[id].getHourData(i);
			vp[i] = lparVP[id].getHourData(i);
			
			cap = lparStatus[id].getHourCap(i);
			ms = lparStatus[id].getHourMS(i);
			pool = lparStatus[id].getHourPool(i);
			
			if (pc[i]<0 || ent[i]<=0) {
				max[i]=-1;
				continue;
			}
			
			if (pool.equals("DefaultPool")) {
				free = msCoreConfig[getMSid(ms)].getHourData(i) - msCoreUsed[getMSid(ms)].getHourData(i);
			} else {
				poolID = getProcPoolId(ms,pool);
				free = procPoolConfig[poolID].getHourData(i)-procPoolUsed[poolID].getHourData(i);
			}
			
			if (cap) {
				max[i] = ent[i];					
			} else {
				if (pc[i]+free>=vp[i])
					max[i] = vp[i];
				else
					max[i] = pc[i]+free;
				if (max[i]<ent[i])
					max[i]=ent[i];
			}
					
		}
		
		
		
		for (i=STEP, j=0; i<24*60; i+=STEP) {
			label[j++] = lparPC[id].getHourLabel(i);
		}
		
		lg = new LparGraph(lparNames[id]+" - 2 Months Hourly", pc, ent, vp, max, label, STEP);
		lg.setSize(XSIZE, YSIZE);
		lg.repaint();
		
		try {
			BufferedImage bi = new BufferedImage(lg.getWidth(), lg.getHeight(), BufferedImage.TYPE_INT_RGB);
			Graphics2D g2d = bi.createGraphics();
			lg.paint(g2d);
				
			try {
				// Save as PNG
				javax.imageio.ImageIO.write(bi, "png", file);
			} catch (IOException e) {
			}
			
		} catch (OutOfMemoryError oome) {
			System.out.println(" Insufficient memory to create PNG output!");
			System.out.println(" Provide more memory to Java using the -Xmx flag.");
		}	
		
	}
	
	
	private void createVirtPoolDailyImage(int id, File file) {
		float pc[] = new float[365];
		float size[] = new float[365];
		float max[] = new float[365];
		
		final int					STEP = 40;
		String label[] = new String[365/STEP];
		int i,j;
		PoolGraph pg=null;
		String poolName = procPoolName[id];	
		String 	msName;
		int		msID;
		
		GregorianCalendar gc;
		
		String split[] = poolName.split("\n");
		msName = split[1];
		for (msID=0; msID<msCoreUsed.length; msID++)
			if (managedSystem[msID].getVarValues("name")[0].equals(msName))
				break;
		
		for (i=0; i<pc.length; i++) {
			pc[i] = procPoolUsed[id].getDayData(i);
			size[i] = procPoolConfig[id].getDayData(i);
			
			if (size[i]<=msCoreConfig[msID].getDayData(i)-msCoreUsed[msID].getDayData(i))
				max[i]=size[i];
			else
				max[i] = msCoreConfig[msID].getDayData(i)-msCoreUsed[msID].getDayData(i);
		}
		
		
		
		for (i=STEP, j=0; i<365; i+=STEP) {
			label[j++]=procPoolUsed[id].getDayLabel(i);
		}
		
		pg = new PoolGraph(split[0]+"@"+split[1]+" - 1 Year Daily", pc, size, max, label, STEP);
		pg.setSize(XSIZE, YSIZE);
		pg.repaint();
		
		try {
			BufferedImage bi = new BufferedImage(pg.getWidth(),pg.getHeight(), BufferedImage.TYPE_INT_RGB);
			Graphics2D g2d = bi.createGraphics();
			pg.paint(g2d);
				
			try {
				// Save as PNG
				javax.imageio.ImageIO.write(bi, "png", file);
			} catch (IOException e) {
			}
			
		} catch (OutOfMemoryError oome) {
			System.out.println(" Insufficient memory to create PNG output!");
			System.out.println(" Provide more memory to Java using the -Xmx flag.");
		}	
		
	}
	
	
	
	
	private void createVirtPoolHourlyImage(int id, File file) {
		float pc[] = new float[24*60];
		float size[] = new float[24*60];
		float max[] = new float[24*60];
		
		final int					STEP = 5*24;
		String label[] = new String[24*60/STEP];
		int i,j;
		PoolGraph pg=null;
		String poolName = procPoolName[id];	
		String 	msName;
		int		msID;
		
		GregorianCalendar gc;
		
		String split[] = poolName.split("\n");
		msName = split[1];
		for (msID=0; msID<msCoreUsed.length; msID++)
			if (managedSystem[msID].getVarValues("name")[0].equals(msName))
				break;
		
		for (i=0; i<pc.length; i++) {
			pc[i] = procPoolUsed[id].getHourData(i); 
			size[i] = procPoolConfig[id].getHourData(i);
			
			if (size[i]<=msCoreConfig[msID].getHourData(i)-msCoreUsed[msID].getHourData(i))
				max[i]=size[i];
			else
				max[i] = msCoreConfig[msID].getHourData(i)-msCoreUsed[msID].getHourData(i);
		}
		
		
		
		for (i=STEP, j=0; i<24*60; i+=STEP) {
			label[j++] =  procPoolUsed[id].getHourLabel(i);
		}
		
		pg = new PoolGraph(split[0]+"@"+split[1]+" - 2 Months Hourly", pc, size, max, label, STEP); 
		pg.setSize(XSIZE, YSIZE);
		pg.repaint();
		
		try {
			BufferedImage bi = new BufferedImage(pg.getWidth(),pg.getHeight(), BufferedImage.TYPE_INT_RGB);
			Graphics2D g2d = bi.createGraphics();
			pg.paint(g2d);
				
			try {
				// Save as PNG
				javax.imageio.ImageIO.write(bi, "png", file);
			} catch (IOException e) {
			}
			
		} catch (OutOfMemoryError oome) {
			System.out.println(" Insufficient memory to create PNG output!");
			System.out.println(" Provide more memory to Java using the -Xmx flag.");
		}	
		
	}
	
	
	
	private void createSysPoolHourlyImage(int id, File file) {
		float pc[] = new float[24*60];
		float size[] = new float[24*60];
		
		final int					STEP = 5*24;
		String label[] = new String[24*60/STEP];
		int i,j;
		PoolGraph pg=null;
		
		GregorianCalendar gc;
		
		for (i=0; i<pc.length; i++) {
			pc[i] = msCoreUsed[id].getHourData(i);
			size[i] = msCoreConfig[id].getHourData(i);
		}
		
		
		
		for (i=STEP, j=0; i<24*60; i+=STEP) {
			label[j++]=msCoreUsed[id].getHourLabel(i);
		}
		
		pg = new PoolGraph(managedSystem[id].getVarValues("name")[0]+" - 2 Months Hourly", pc, size, label, STEP);
		pg.setSize(XSIZE, YSIZE);
		pg.repaint();
		
		try {
			BufferedImage bi = new BufferedImage(pg.getWidth(),pg.getHeight(), BufferedImage.TYPE_INT_RGB);
			Graphics2D g2d = bi.createGraphics();
			pg.paint(g2d);
				
			try {
				// Save as PNG
				javax.imageio.ImageIO.write(bi, "png", file);
			} catch (IOException e) {
			}
			
		} catch (OutOfMemoryError oome) {
			System.out.println(" Insufficient memory to create PNG output!");
			System.out.println(" Provide more memory to Java using the -Xmx flag.");
		}	
		
	}
	
	
	
	
	private void createSysPoolDailyImage(int id, File file) {
		float pc[] = new float[365];
		float size[] = new float[365];
		
		final int					STEP = 40;
		String label[] = new String[365/STEP];
		int i,j;
		PoolGraph pg=null;
		
		GregorianCalendar gc;
		
		for (i=0; i<pc.length; i++) {
			pc[i] = msCoreUsed[id].getDayData(i);
			size[i] = msCoreConfig[id].getDayData(i);					
		}
		
		
		
		for (i=STEP, j=0; i<365; i+=STEP) {
			label[j++]=msCoreUsed[id].getDayLabel(i);
		}
		
		pg = new PoolGraph(managedSystem[id].getVarValues("name")[0]+" - 1 Year Daily", pc, size, label, STEP);
		pg.setSize(XSIZE, YSIZE);
		pg.repaint();
		
		try {
			BufferedImage bi = new BufferedImage(pg.getWidth(), pg.getHeight(), BufferedImage.TYPE_INT_RGB);
			Graphics2D g2d = bi.createGraphics();
			pg.paint(g2d);
				
			try {
				// Save as PNG
				javax.imageio.ImageIO.write(bi, "png", file);
			} catch (IOException e) {
			}
			
		} catch (OutOfMemoryError oome) {
			System.out.println(" Insufficient memory to create PNG output!");
			System.out.println(" Provide more memory to Java using the -Xmx flag.");
		}	
		
	}
	
	
	
	
	private void createLparDailyImage(int id, File file) {
		float pc[] = new float[365];
		float ent[] = new float[365];
		float vp[] = new float[365];
		float max[] = new float[365];
		
		final int					STEP = 40;
		String label[] = new String[365/STEP];
		int i,j;
		LparGraph lg=null;
		
		boolean cap;
		String ms;
		String pool;
		float free;
		int poolID;
		GregorianCalendar gc;
		
		for (i=0; i<pc.length; i++) {
			pc[i] = lparPC[id].getDayData(i);
			ent[i] = lparEnt[id].getDayData(i);
			vp[i] = lparVP[id].getDayData(i);
			
			cap = lparStatus[id].getDayCap(i);
			ms = lparStatus[id].getDayMS(i);
			pool = lparStatus[id].getDayPool(i);
			
			if (pc[i]<0 || ent[i]<=0) {
				max[i]=-1;
				continue;
			}
			
			if (pool.equals("DefaultPool")) {
				free = msCoreConfig[getMSid(ms)].getDayData(i) - msCoreUsed[getMSid(ms)].getDayData(i);
			} else {
				poolID = getProcPoolId(ms,pool);
				free = procPoolConfig[poolID].getDayData(i)-procPoolUsed[poolID].getDayData(i);
			}
			
			if (cap) {
				max[i] = ent[i];					
			} else {
				if (pc[i]+free>=vp[i])
					max[i] = vp[i];
				else
					max[i] = pc[i]+free;
				if (max[i]<ent[i])
					max[i]=ent[i];
			}
					
		}
		
		
		
		for (i=STEP, j=0; i<365; i+=STEP) {
			label[j++]=lparPC[id].getDayLabel(i);
		}
		
		lg = new LparGraph(lparNames[id]+" - 1 Year Daily", pc, ent, vp, max, label, STEP);
		lg.setSize(XSIZE, YSIZE);
		lg.repaint();
		
		try {
			BufferedImage bi = new BufferedImage(lg.getWidth(), lg.getHeight(), BufferedImage.TYPE_INT_RGB);
			Graphics2D g2d = bi.createGraphics();
			lg.paint(g2d);
				
			try {
				// Save as PNG
				javax.imageio.ImageIO.write(bi, "png", file);
			} catch (IOException e) {
			}
			
		} catch (OutOfMemoryError oome) {
			System.out.println(" Insufficient memory to create PNG output!");
			System.out.println(" Provide more memory to Java using the -Xmx flag.");
		}	
		
	}
	
	
	
	private void createIOChildrenSheetExcel(WritableSheet sheet) {
		DataSheet ds = createIOChildrenSheet();
		ds.createExcelSheet(sheet);
	}
	
	private void createIOChildrenSheetHTML(String fileName) {
		DataSheet ds = createIOChildrenSheet();
		ds.createHTMLSheet(fileName);
		addButton("IO Children",new File(fileName).getName());
	}
	
	private void createIOChildrenSheetCSV(String fileName) {
		DataSheet ds = createIOChildrenSheet();
		ds.setSeparator(csvSeparator);
		ds.createCSVSheet(fileName);
	}
	
	
	private DataSheet createIOChildrenSheet() {
		DataSheet ds = new DataSheet();
		GenericData children[];
		int row,col;
		int i,j;
		String s[];
		int size[]=new int[7];
		int n;
			
		row = 0;
		col=0;
		for (i=0; i<size.length; i++)
			size[i] = 0;
		
		
		/*
		 * Setup titles
		 */ 

		n = ds.addLabel(col,row,"Physical Location",BOLD|VCENTRE|B_ALL_MED|GREEN);	if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Description",BOLD|VCENTRE|B_ALL_MED|GREEN);if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Microcode Version",BOLD|VCENTRE|B_ALL_MED|GREEN);if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"MAC Address",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"WWPN",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;		
		
		n = ds.addLabel(col,row,"Managed System Name",BOLD|CENTRE|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Managed System Serial",BOLD|CENTRE|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		
		row++;
		
		

		for (i=0; i<managedSystem.length; i++) {
			/*
			 * Write variables
			 */
			children = managedSystem[i].getObjects(IOSLOTCHILDREN);
			if (children==null) 
				continue;
			
			for (j=0; j<children.length; j++) {
				
				col=0;
				
				s = children[j].getVarValues("phys_loc");
				n = ds.addLabel(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				s = children[j].getVarValues("description");
				n = ds.addLabel( col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				s = children[j].getVarValues("microcode_version");
				n = ds.addLabel( col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				s = children[j].getVarValues("mac_address");
				n = ds.addLabel( col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				s = children[j].getVarValues("wwpn");
				n = ds.addLabel( col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
					
				
				s = managedSystem[i].getVarValues("name");
				n = ds.addLabel( col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				s = managedSystem[i].getVarValues("serial_num");
				n = ds.addLabel( col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
							
				row++;				
			}
		}
		
		for (i=0; i<size.length; i++)
			ds.setColSize(i, size[i]+2);

		return ds;
	}
	
	
	private DataSheet createOnOffSheet() {
		DataSheet ds = new DataSheet();
		GenericData onoff_cap[];
		int row,col;
		int i,j;
		String s[];
		int size[]=new int[15];
		int n;
			
		row = 0;
		col=0;
		for (i=0; i<size.length; i++)
			size[i] = 0;
		
		
		/*
		 * Setup titles
		 */ 

		n = ds.addLabel(col,row,"Type",BOLD|VCENTRE|B_ALL_MED|GREEN);	if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"State",BOLD|VCENTRE|B_ALL_MED|GREEN);if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Activated",BOLD|VCENTRE|B_ALL_MED|GREEN);if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Available",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Unreturned",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Days left",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Hours left",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Days available",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Report date",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Report time",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Hist expired days",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Hist unreturned days",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Tot sys runtime hours",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		
		n = ds.addLabel(col,row,"Managed System Name",BOLD|CENTRE|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Managed System Serial",BOLD|CENTRE|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		
		row++;
		
		

		for (i=0; i<managedSystem.length; i++) {
			/*
			 * Write variables
			 */
			onoff_cap = managedSystem[i].getObjects(LSCOD_CAP_PROC_ONOFF);
			if (onoff_cap!=null && onoff_cap.length>0) {
					
				col=0;
				
				n = ds.addLabel(col, row, "CPU", B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				s = onoff_cap[0].getVarValues("proc_onoff_state");
				n = ds.addLabel(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				s = onoff_cap[0].getVarValues("activated_onoff_procs");
				ds.addInteger(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); col++;
			
				s = onoff_cap[0].getVarValues("avail_procs_for_onoff");
				ds.addInteger(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); col++;	
				
				s = onoff_cap[0].getVarValues("unreturned_onoff_procs");
				ds.addInteger(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); col++;
				
				s = onoff_cap[0].getVarValues("onoff_request_proc_days_left");
				ds.addInteger(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); col++;
				
				s = onoff_cap[0].getVarValues("onoff_proc_day_hours_left");
				ds.addInteger(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); col++;
				
				s = onoff_cap[0].getVarValues("onoff_proc_days_avail");
				ds.addInteger(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); col++;
				
				
				onoff_cap = managedSystem[i].getObjects(LSCOD_BILL_PROC);
				if (onoff_cap!=null && onoff_cap.length>0) {
															
					s = onoff_cap[0].getVarValues("collection_date");
					n = ds.addLabel(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
					
					s = onoff_cap[0].getVarValues("collection_time");
					n = ds.addLabel(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
					
					s = onoff_cap[0].getVarValues("hist_expired_resource_days");
					ds.addInteger(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); col++;
					
					s = onoff_cap[0].getVarValues("hist_unreturned_resource_days");
					ds.addInteger(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); col++;
					
					s = onoff_cap[0].getVarValues("total_sys_run_time_hours");
					ds.addInteger(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); col++;
								
				}
				
				
				s = managedSystem[i].getVarValues("name");
				n = ds.addLabel( col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				s = managedSystem[i].getVarValues("serial_num");
				n = ds.addLabel( col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				
				row++;
			}
			
			
			
			
			onoff_cap = managedSystem[i].getObjects(LSCOD_CAP_MEM_ONOFF);
			if (onoff_cap!=null && onoff_cap.length>0) {
					
				col=0;
				
				n = ds.addLabel(col, row, "MEM", B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				s = onoff_cap[0].getVarValues("mem_onoff_state");
				n = ds.addLabel(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				s = onoff_cap[0].getVarValues("activated_onoff_mem");
				ds.addInteger(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); col++;
			
				s = onoff_cap[0].getVarValues("avail_mem_for_onoff");
				ds.addInteger(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); col++;	
				
				s = onoff_cap[0].getVarValues("unreturned_onoff_mem");
				ds.addInteger(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); col++;
				
				s = onoff_cap[0].getVarValues("onoff_request_mem_days_left");
				ds.addInteger(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); col++;
				
				s = onoff_cap[0].getVarValues("onoff_mem_day_hours_left");
				ds.addInteger(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); col++;
				
				s = onoff_cap[0].getVarValues("onoff_mem_days_avail");
				ds.addInteger(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); col++;
				
				
				onoff_cap = managedSystem[i].getObjects(LSCOD_BILL_MEM);
				if (onoff_cap!=null && onoff_cap.length>0) {
															
					s = onoff_cap[0].getVarValues("collection_date");
					n = ds.addLabel(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
					
					s = onoff_cap[0].getVarValues("collection_time");
					n = ds.addLabel(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
					
					s = onoff_cap[0].getVarValues("hist_expired_resource_days");
					ds.addInteger(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); col++;
					
					s = onoff_cap[0].getVarValues("hist_unreturned_resource_days");
					ds.addInteger(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); col++;
					
					s = onoff_cap[0].getVarValues("total_sys_run_time_hours");
					ds.addInteger(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); col++;
								
				}
				
				
				s = managedSystem[i].getVarValues("name");
				n = ds.addLabel( col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				s = managedSystem[i].getVarValues("serial_num");
				n = ds.addLabel( col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
							
				row++;
			}
			
					
		}
		
		for (i=0; i<size.length; i++)
			ds.setColSize(i, size[i]+2);

		return ds;
	}
	
	
	private void createOnOffSheetExcel(WritableSheet sheet) {
		DataSheet ds = createOnOffSheet();
		ds.createExcelSheet(sheet);
	}
	
	private void createOnOffSheetHTML(String fileName) {
		DataSheet ds = createOnOffSheet();
		ds.createHTMLSheet(fileName);
		addButton("OnOff",new File(fileName).getName());
	}
	
	private void createOnOffSheetCSV(String fileName) {
		DataSheet ds = createOnOffSheet();
		ds.setSeparator(csvSeparator);
		ds.createCSVSheet(fileName);
	}
	
	
	private DataSheet createCoDLogSheet() {
		DataSheet ds = new DataSheet();
		GenericData event[];
		int row,col;
		int i,j;
		String s[];
		int size[]=new int[4]; 
		int n;
		String date;
			
		row = 0;
		col=0;
		for (i=0; i<size.length; i++)
			size[i] = 0;
		
		
		/*
		 * Setup titles
		 */ 

		n = ds.addLabel(col,row,"Time",BOLD|VCENTRE|B_ALL_MED|GREEN);	if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Event",BOLD|VCENTRE|B_ALL_MED|GREEN);if (n>size[col]) size[col]=n; col++;
		
		n = ds.addLabel(col,row,"Managed System Name",BOLD|CENTRE|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Managed System Serial",BOLD|CENTRE|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		
		row++;
		
		

		for (i=0; i<managedSystem.length; i++) {
			/*
			 * Write variables
			 */
			event = managedSystem[i].getObjects(LSCOD_HIST);
			if (event==null)
				continue;
			for (j=event.length-1; j>=0; j--) {
					
				col=0;
								
				s = event[j].getVarValues("time_stamp");
				
				// Time stamp is provided is silly US format. Change it to a sortable format: YYYY/MM/DD HH:MM:SS
				// 012345678901234567890
				// 04/28/2014 08:43:23
				date = s[0].substring(6, 10)+"/"+s[0].substring(0, 2)+"/"+s[0].substring(3, 5)+" "+
						s[0].substring(11);
				
				n = ds.addLabel(col, row, date, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				s = event[j].getVarValues("entry");
				n = ds.addLabel(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;

							
				s = managedSystem[i].getVarValues("name");
				n = ds.addLabel( col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				s = managedSystem[i].getVarValues("serial_num");
				n = ds.addLabel( col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				
				row++;
			}
					
		}
		
		for (i=0; i<size.length; i++)
			ds.setColSize(i, size[i]+2);

		return ds;
	}
	
	
	private void createCoDLogSheetExcel(WritableSheet sheet) {
		DataSheet ds = createCoDLogSheet();
		ds.createExcelSheet(sheet);
	}
	
	private void createCoDLogSheetHTML(String fileName) {
		DataSheet ds = createCoDLogSheet();
		ds.createHTMLSheet(fileName);
		addButton("OnOff Log",new File(fileName).getName());
	}
	
	private void createCoDLogSheetCSV(String fileName) {
		DataSheet ds = createCoDLogSheet();
		ds.setSeparator(csvSeparator);
		ds.createCSVSheet(fileName);
	}
	
	
	private void createIOChildrenSheet(WritableSheet sheet)  throws RowsExceededException, WriteException {
		GenericData children[];
		int row,col;
		int i,j;
		String s[];
		int size[]=new int[7];
		int n;
			
		row = 0;
		col=0;
		for (i=0; i<size.length; i++)
			size[i] = 0;
		
		
		/*
		 * Setup titles
		 */ 

		n = addLabel(sheet,col,row,"Physical Location",formatLabel(BOLD|VCENTRE|B_ALL_MED|GREEN));	if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"Description",formatLabel(BOLD|VCENTRE|B_ALL_MED|GREEN));if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"Microcode Version",formatLabel(BOLD|VCENTRE|B_ALL_MED|GREEN));if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"MAC Address",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"WWPN",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;		
		
		n = addLabel(sheet,col,row,"Managed System Name",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"Managed System Serial",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		
		row++;
		
		

		for (i=0; i<managedSystem.length; i++) {
			/*
			 * Write variables
			 */
			children = managedSystem[i].getObjects(IOSLOTCHILDREN);
			if (children==null) 
				continue;
			
			for (j=0; j<children.length; j++) {
				
				col=0;
				
				s = children[j].getVarValues("phys_loc");
				n = addLabel(sheet, col, row, s, 0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				if (n>size[col]) size[col]=n; col++;
				
				s = children[j].getVarValues("description");
				n = addLabel(sheet, col, row, s, 0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				if (n>size[col]) size[col]=n; col++;
				
				s = children[j].getVarValues("microcode_version");
				n = addLabel(sheet, col, row, s, 0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				if (n>size[col]) size[col]=n; col++;
				
				s = children[j].getVarValues("mac_address");
				n = addLabel(sheet, col, row, s, 0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				if (n>size[col]) size[col]=n; col++;
				
				s = children[j].getVarValues("wwpn");
				n = addLabel(sheet, col, row, s, 0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				if (n>size[col]) size[col]=n; col++;
					
				
				s = managedSystem[i].getVarValues("name");
				n = addLabel(sheet, col, row, s, 0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				if (n>size[col]) size[col]=n; col++;
				
				s = managedSystem[i].getVarValues("serial_num");
				n = addLabel(sheet, col, row, s, 0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				if (n>size[col]) size[col]=n; col++;
				
							
				row++;				
			}
		}
		
		for (i=0; i<size.length; i++)
			sheet.setColumnView(i, size[i]+2);
	}
	
	
	
	private void createProfileSheetExcel(WritableSheet sheet) {
		DataSheet ds = createProfileSheet();
		if (ds!=null)
			ds.createExcelSheet(sheet);
	}
	
	private void createProfileSheetHTML(String fileName) {
		DataSheet ds = createProfileSheet();
		if (ds!=null) {
			ds.createHTMLSheet(fileName);
			//addButton("LPAR Profiles",profile_html);
			addButton("LPAR Profiles",new File(fileName).getName());
		}
	}
	
	private void createProfileSheetCSV(String fileName) {
		DataSheet ds = createProfileSheet();
		if (ds!=null) {
			ds.setSeparator(csvSeparator);
			ds.createCSVSheet(fileName);
		}
	}
	
	
	
	private DataSheet createProfileSheet() {
		DataSheet ds = new DataSheet();
		GenericData profile[];
		int row,col;
		int i,j;
		String s[];
		String activeProfile;
		boolean active;
		int size[]=new int[35];
		int n;
			
		row = 0;
		col=0;
		for (i=0; i<size.length; i++)
			size[i] = 0;
		
		
		/*
		 * Setup titles
		 */ 

		n = ds.addLabel(col,row,"Lpar Name",BOLD|VCENTRE|B_ALL_MED|GREEN);	if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Profile Name",BOLD|VCENTRE|B_ALL_MED|GREEN);if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Active",BOLD|VCENTRE|B_ALL_MED|GREEN);if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"ID",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Env",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"All Res",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"MinMem",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"DesMem",BOLD|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"MaxMem",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"MinHugePages",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"DesHugePages",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"MaxHugePages",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"MemMode",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"MemExp",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"HPT_ratio",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"ProcMode",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"MinEnt",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"DesEnt",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"MaxEnt",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"MinProc",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"DesProc",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Maxproc",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"ShMode",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Weight",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"ProcPoolID",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"ProcPoolName",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"AffinityGrp",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"BootMode",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"ConnMonit",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"AutoStart",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"BSR_arrays",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"ProcCompat",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"ErrReporting",BOLD|RIGHT|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;

		
		
		
		n = ds.addLabel(col,row,"Managed System Name",BOLD|CENTRE|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"Managed System Serial",BOLD|CENTRE|B_ALL_MED|GREEN); if (n>size[col]) size[col]=n; col++;
		
		row++;
		
		

		for (i=0; i<managedSystem.length; i++) {
			/*
			 * Write variables
			 */
			profile = managedSystem[i].getObjects(PROFILES);
			if (profile==null) 
				continue;
			
			for (j=0; j<profile.length; j++) {
				
				col=0;
				
				s = profile[j].getVarValues("lpar_name");
				n = ds.addLabel(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				activeProfile = getActiveProfileName(i,s[0]);
				
				s = profile[j].getVarValues("name");
				if (activeProfile!=null && s!=null && activeProfile.equals(s[0]))
					active=true;
				else
					active=false;
					
				n = ds.addLabel(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				if (active)
					n = ds.addLabel(col, row, "true", B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
				else
					n = ds.addLabel(col, row, "false", B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
				col++;

				
				s = profile[j].getVarValues("lpar_id");
				ds.addInteger(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);	
				n=4; if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("lpar_env");
				n = ds.addLabel(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
				if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("all_resources");
				if (s[0].equals("0"))
					n = ds.addLabel(col, row, "false", B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
				else
					n = ds.addLabel(col, row, "true", B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
				if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("min_mem");
				ds.addFloatDiv1024(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
				n=8; if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("desired_mem");
				ds.addFloatDiv1024(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
				n=8; if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("max_mem");
				ds.addFloatDiv1024(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
				n=8; if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("min_num_huge_pages");
				ds.addInteger(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
				n=8; if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("desired_num_huge_pages");
				ds.addInteger(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
				n=8; if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("max_num_huge_pages");
				ds.addInteger(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
				n=8; if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("mem_mode");
				n = ds.addLabel(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("mem_expansion");
				ds.addFloat(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
				n=8; if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("hpt_ratio");
				n = ds.addLabel(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("proc_mode");
				n = ds.addLabel(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("min_proc_units");
				ds.addFloat(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
				n=6; if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("desired_proc_units");
				ds.addFloat(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
				n=6; if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("max_proc_units");
				ds.addFloat(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
				n=6; if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("min_procs");
				ds.addInteger(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
				n=3; if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("desired_procs");
				ds.addInteger(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
				n=3; if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("max_procs");
				ds.addInteger(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
				n=3; if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("sharing_mode");
				n = ds.addLabel(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("uncap_weight");
				ds.addInteger(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
				n=3; if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("shared_proc_pool_id");
				ds.addInteger(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
				n=3; if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("shared_proc_pool_name");
				n = ds.addLabel(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("affinity_group_id");
				n = ds.addLabel(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("boot_mode");
				n = ds.addLabel(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("conn_monitoring");
				n = ds.addLabel(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("auto_start");
				n = ds.addLabel(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("bsr_arrays");
				n = ds.addLabel(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("lpar_proc_compat_mode");
				n = ds.addLabel(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("electronic_err_reporting");
				n = ds.addLabel(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				

				
				
				s = managedSystem[i].getVarValues("name");
				n = ds.addLabel(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
				s = managedSystem[i].getVarValues("serial_num");
				n = ds.addLabel(col, row, s, 0, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW); if (n>size[col]) size[col]=n; col++;
				
							
				row++;				
			}
		}
		
		for (i=0; i<size.length; i++)
			ds.setColSize(i, size[i]+2);

		return ds;
	}
	
	
	
	
	private void createProfileSheet(WritableSheet sheet)  throws RowsExceededException, WriteException {
		GenericData profile[];
		int row,col;
		int i,j;
		String s[];
		String activeProfile;
		boolean active;
		double d;	
		int size[]=new int[35];
		int n;
			
		row = 0;
		col=0;
		for (i=0; i<size.length; i++)
			size[i] = 0;
		
		
		/*
		 * Setup titles
		 */ 

		n = addLabel(sheet,col,row,"Lpar Name",formatLabel(BOLD|VCENTRE|B_ALL_MED|GREEN));	if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"Profile Name",formatLabel(BOLD|VCENTRE|B_ALL_MED|GREEN));if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"Active",formatLabel(BOLD|VCENTRE|B_ALL_MED|GREEN));if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"ID",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"Env",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"All Res",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"MinMem",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"DesMem",formatLabel(BOLD|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"MaxMem",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"MinHugePages",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"DesHugePages",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"MaxHugePages",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"MemMode",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"MemExp",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"HPT_ratio",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"ProcMode",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"MinEnt",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"DesEnt",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"MaxEnt",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"MinProc",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"DesProc",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"Maxproc",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"ShMode",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"Weight",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"ProcPoolID",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"ProcPoolName",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"AffinityGrp",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"BootMode",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"ConnMonit",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"AutoStart",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"BSR_arrays",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"ProcCompat",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"ErrReporting",formatLabel(BOLD|RIGHT|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;

		
		
		
		n = addLabel(sheet,col,row,"Managed System Name",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"Managed System Serial",formatLabel(BOLD|CENTRE|B_ALL_MED|GREEN)); if (n>size[col]) size[col]=n; col++;
		
		row++;
		
		

		for (i=0; i<managedSystem.length; i++) {
			/*
			 * Write variables
			 */
			profile = managedSystem[i].getObjects(PROFILES);
			if (profile==null) 
				continue;
			
			for (j=0; j<profile.length; j++) {
				
				col=0;
				
				s = profile[j].getVarValues("lpar_name");
				n = addLabel(sheet, col, row, s, 0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				if (n>size[col]) size[col]=n; col++;
				
				activeProfile = getActiveProfileName(i,s[0]);
				
				s = profile[j].getVarValues("name");
				if (activeProfile!=null && s!=null && activeProfile.equals(s[0]))
					active=true;
				else
					active=false;
					
				n = addLabel(sheet, col, row, s, 0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				if (n>size[col]) size[col]=n; col++;
				
				if (active)
					addLabel(sheet, col, row, "true", formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				else
					addLabel(sheet, col, row, "false", formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				col++;

				
				s = profile[j].getVarValues("lpar_id");
				addNumber(sheet, col, row, s, 0, formatInt(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));	
				n=4;
				if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("lpar_env");
				n = addLabel(sheet, col, row, s, 0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("all_resources");
				if (s[0].equals("0"))
					n = addLabel(sheet, col, row, "false", formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				else
					n = addLabel(sheet, col, row, "true", formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("min_mem");
				addNumberDiv1024(sheet, col, row, s, 0, formatFloat(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				n=8;
				if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("desired_mem");
				addNumberDiv1024(sheet, col, row, s, 0, formatFloat(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				n=8;
				if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("max_mem");
				addNumberDiv1024(sheet, col, row, s, 0, formatFloat(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				n=8;
				if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("min_num_huge_pages");
				addNumber(sheet, col, row, s, 0, formatInt(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				n=8;
				if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("desired_num_huge_pages");
				addNumber(sheet, col, row, s, 0, formatInt(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				n=8;
				if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("max_num_huge_pages");
				addNumber(sheet, col, row, s, 0, formatInt(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				n=8;
				if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("mem_mode");
				n = addLabel(sheet, col, row, s, 0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("mem_expansion");
				addNumber(sheet, col, row, s, 0, formatFloat(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				n=8;
				if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("hpt_ratio");
				n = addLabel(sheet, col, row, s, 0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("proc_mode");
				n = addLabel(sheet, col, row, s, 0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("min_proc_units");
				addNumber(sheet, col, row, s, 0, formatFloat(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				n=6;
				if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("desired_proc_units");
				addNumber(sheet, col, row, s, 0, formatFloat(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				n=6;
				if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("max_proc_units");
				addNumber(sheet, col, row, s, 0, formatFloat(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				n=6;
				if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("min_procs");
				addNumber(sheet, col, row, s, 0, formatInt(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				n=3;
				if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("desired_procs");
				addNumber(sheet, col, row, s, 0, formatInt(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				n=3;
				if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("max_procs");
				addNumber(sheet, col, row, s, 0, formatInt(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				n=3;
				if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("sharing_mode");
				n = addLabel(sheet, col, row, s, 0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("uncap_weight");
				addNumber(sheet, col, row, s, 0, formatInt(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				n=3;
				if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("shared_proc_pool_id");
				addNumber(sheet, col, row, s, 0, formatInt(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				n=3;
				if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("shared_proc_pool_name");
				n = addLabel(sheet, col, row, s, 0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("affinity_group_id");
				n = addLabel(sheet, col, row, s, 0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("boot_mode");
				n = addLabel(sheet, col, row, s, 0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("conn_monitoring");
				n = addLabel(sheet, col, row, s, 0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("auto_start");
				n = addLabel(sheet, col, row, s, 0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("bsr_arrays");
				n = addLabel(sheet, col, row, s, 0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("lpar_proc_compat_mode");
				n = addLabel(sheet, col, row, s, 0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				if (n>size[col]) size[col]=n; col++;
				
				s = profile[j].getVarValues("electronic_err_reporting");
				n = addLabel(sheet, col, row, s, 0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				if (n>size[col]) size[col]=n; col++;
				

				
				
				s = managedSystem[i].getVarValues("name");
				n = addLabel(sheet, col, row, s, 0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				if (n>size[col]) size[col]=n; col++;
				
				s = managedSystem[i].getVarValues("serial_num");
				n = addLabel(sheet, col, row, s, 0, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				if (n>size[col]) size[col]=n; col++;
				
							
				row++;				
			}
		}
		
		for (i=0; i<size.length; i++)
			sheet.setColumnView(i, size[i]+2);
	}
	
	
	private void createViosDiskSheetExcel(WritableSheet sheet) {
		DataSheet ds = createViosDiskSheet();
		if (ds!=null)
			ds.createExcelSheet(sheet);
	}
	
	private void createViosDiskSheetHTML(String fileName) {
		DataSheet ds = createViosDiskSheet();
		if (ds!=null) {
			ds.createHTMLSheet(fileName);
			addButton("VIOS disks",new File(fileName).getName());
		}
	}
	
	private void createViosDiskSheetCSV(String fileName) {
		DataSheet ds = createViosDiskSheet();
		if (ds!=null) {
			ds.setSeparator(csvSeparator);
			ds.createCSVSheet(fileName);
		}
	}
	
	private DataSheet createViosDiskSheet() {
		DataSheet ds = new DataSheet();
		int row,col;
		int i,j;
		int n;

		
		/*
		 * Setup titles
		 */ 		
		String vios[] = diskData.getViosNames();
		int size[]=new int[3+vios.length];
		row = 0;
		col=0;
		for (i=0; i<size.length; i++)
			size[i] = 0;

		n = ds.addLabel(col,row,"UUID",BOLD|VCENTRE|B_ALL_MED|GREEN);if (n>size[col]) size[col]=n; col++;
		n = ds.addLabel(col,row,"SIZE",BOLD|VCENTRE|B_ALL_MED|GREEN);if (n>size[col]) size[col]=n; col++;
		
		for (i=0; i<vios.length; i++) {
			n = ds.addLabel(col,row,vios[i],BOLD|VCENTRE|B_ALL_MED|GREEN);
			if (n>size[col]) size[col]=n; 
			col++;
		}
	
		row++;
		
		
		int num_uuid = diskData.getNumUUID();	
		String hdiskName[];
		boolean free[];
		String uuidName;
		int    diskSize;
		int map;

		for (i=0; i<num_uuid; i++) {
			/*
			 * Write variables
			 */
			
			col = 0;
			
			uuidName = diskData.getUUIDname(i);
			diskSize = diskData.getSize(i);
			hdiskName = diskData.getHdiskOnViosNames(i);
			free = diskData.getFreeOnViosNames(i);
			
			n = ds.addLabel(col, row, uuidName, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
			if (n>size[col]) size[col]=n; col++;
			
			if (diskSize>0) {
				ds.addInteger(col, row, diskSize, B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW);
				n=8; if (n>size[col]) size[col]=n; col++;
			} else {
				col++;
			}
			
			for (j=0; j<hdiskName.length; j++) {
				map = B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW;
				if (free[j])
					map = map | GREEN;
				n = ds.addLabel(col, row, hdiskName, j, map);
				if (n>size[col]) size[col]=n; col++;
			}
			
	
			row++;				
		}
		
		for (i=0; i<size.length; i++)
			ds.setColSize(i, size[i]+2);

		return ds;
	}
	
	
	private void createViosDiskSheet(WritableSheet sheet)  throws RowsExceededException, WriteException {
		int row,col;
		int i,j;
		int n;
			

		
		/*
		 * Setup titles
		 */ 
		
		String vios[] = diskData.getViosNames();
		int size[]=new int[3+vios.length];
		row = 0;
		col=0;
		for (i=0; i<size.length; i++)
			size[i] = 0;

		n = addLabel(sheet,col,row,"UUID",formatLabel(BOLD|VCENTRE|B_ALL_MED|GREEN));if (n>size[col]) size[col]=n; col++;
		n = addLabel(sheet,col,row,"SIZE",formatLabel(BOLD|VCENTRE|B_ALL_MED|GREEN));if (n>size[col]) size[col]=n; col++;
		
		for (i=0; i<vios.length; i++) {
			n = addLabel(sheet,col,row,vios[i],formatLabel(BOLD|VCENTRE|B_ALL_MED|GREEN));
			if (n>size[col]) size[col]=n; 
			col++;
		}
	
		row++;
		
		
		int num_uuid = diskData.getNumUUID();	
		String hdiskName[];
		boolean free[];
		String uuidName;
		int    diskSize;
		int map;

		for (i=0; i<num_uuid; i++) {
			/*
			 * Write variables
			 */
			
			col = 0;
			
			uuidName = diskData.getUUIDname(i);
			diskSize = diskData.getSize(i);
			hdiskName = diskData.getHdiskOnViosNames(i);
			free = diskData.getFreeOnViosNames(i);
			
			n = addLabel(sheet, col, row, uuidName, formatLabel(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
			if (n>size[col]) size[col]=n; col++;
			
			if (diskSize>0) {
				addNumber(sheet, col, row, diskSize, formatInt(B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW));
				size[col]=10; col++;
			} else {
				col++;
			}
			
			for (j=0; j<hdiskName.length; j++) {
				map = B_LEFT_LOW|B_BOTTOM_LOW|B_RIGHT_LOW;
				if (free[j])
					map = map | GREEN;
				n = addLabel(sheet, col, row, hdiskName, j, formatLabel(map));
				if (n>size[col]) size[col]=n; col++;
			}
			
	
			row++;				
		}
		
		for (i=0; i<size.length; i++)
			sheet.setColumnView(i, size[i]+2);
	}
	
	
	/*

	private void getSystemData() {
		BufferedReader br;
		int i,j,k,x;
		int lines;
		String s[];
		
		String system;
		DataParser dp;
		String names[]=null;
		GenericData gd=null;
		
		
		// ===>>>     -F --header !!!!!!!
		
		
		// Get system names and feature configuration
		lines = sshm.sendCommand("lssyscfg -r sys", baseDir + systemData);

				
		managedSystem = new GenericData[lines];
		for (i=0; i<lines; i++)
			managedSystem[i] = new GenericData();
		
		try {			
			br = new BufferedReader(new FileReader(baseDir + systemData),1024*1024);
			

			names=null;
			for (i=0; i<lines; i++) {
				system = br.readLine();
				
				// Skip line if no data is returned
				if (system.startsWith("HSC"))
					continue;
				
				dp = new DataParser(system);
				names = dp.getNames();
				
				for (j=0; j<names.length; j++) {
					managedSystem[i].add(names[j], dp.getStringValue(names[j]));					
				}				
			}			
		} 
		catch (IOException ioe) {	
			System.out.println("Loader.getSystemData: IOException");
			System.out.println(ioe);
		}
		
		
		// Get CPU configuration of systems
		for (i=0; i<managedSystem.length; i++) {
			s = managedSystem[i].getVarValues("name");
			lines = sshm.sendCommand("lshwres -r proc --level sys -m "+s[0], baseDir + s[0] + "_" + procSysData);
			
			try {			
				br = new BufferedReader(new FileReader(baseDir + s[0] + "_" + procSysData),1024*1024);
				
				names=null;
				for (j=0; j<lines; j++) {
					system = br.readLine();
					
					// Skip line if no data is returned
					if (system.startsWith("HSC"))
						continue;
					
					dp = new DataParser(system);
					names = dp.getNames();
					
					for (k=0; k<names.length; k++) {
						managedSystem[i].add(names[k], dp.getStringValue(names[k]));					
					}				
				}			
			} 
			catch (IOException ioe) {	
				System.out.println("Loader.getSystemData: IOException");
				System.out.println(ioe);
			}			
		}
		
		// Get MEM configuration of systems
		for (i=0; i<managedSystem.length; i++) {
			s = managedSystem[i].getVarValues("name");
			lines = sshm.sendCommand("lshwres -r mem --level sys -m "+s[0], baseDir + s[0] + "_" + memSysData);
			
			try {			
				br = new BufferedReader(new FileReader(baseDir + s[0] + "_" + memSysData),1024*1024);
				
				names=null;
				for (j=0; j<lines; j++) {
					system = br.readLine();
					
					// Skip line if no data is returned
					if (system.startsWith("HSC"))
						continue;
					
					dp = new DataParser(system);
					names = dp.getNames();
					
					for (k=0; k<names.length; k++) {
						managedSystem[i].add(names[k], dp.getStringValue(names[k]));					
					}				
				}			
			} 
			catch (IOException ioe) {	
				System.out.println("Loader.getSystemData: IOException");
				System.out.println(ioe);
			}			
		}
		
		// Get SLOT configuration of systems
		for (i=0; i<managedSystem.length; i++) {
			s = managedSystem[i].getVarValues("name");
			lines = sshm.sendCommand("lshwres -r io --rsubtype slot -m "+s[0], baseDir + s[0] + "_" + slotSysData);
			
			try {			
				br = new BufferedReader(new FileReader(baseDir + s[0] + "_" + slotSysData),1024*1024);
				
				names=null;
				
				for (j=0; j<lines; j++) {
					system = br.readLine();
					
					// Skip line if no data is returned
					if (system.startsWith("HSC"))
						continue;
					
					dp = new DataParser(system);
					names = dp.getNames();
					
					gd = new GenericData();
					
					for (k=0; k<names.length; k++) {
						gd.add(names[k], dp.getStringValue(names[k]));					
					}	
					
					managedSystem[i].addObject("slot", gd);
				}			
			} 
			catch (IOException ioe) {	
				System.out.println("Loader.getSystemData: IOException");
				System.out.println(ioe);
			}			
		}
		
		
		
		// Get CPU configuration of LPARs
		for (i=0; i<managedSystem.length; i++) {
			s = managedSystem[i].getVarValues("name");
			lines = sshm.sendCommand("lshwres -r proc --level lpar -m "+s[0], baseDir + s[0] + "_" + procLparData);
			
			try {			
				br = new BufferedReader(new FileReader(baseDir + s[0] + "_" + procLparData),1024*1024);
				
				names=null;
				
				for (j=0; j<lines; j++) {
					system = br.readLine();
					
					// Skip line if no data is returned
					if (system.startsWith("HSC"))
						continue;
					
					dp = new DataParser(system);
					names = dp.getNames();
					
					gd = new GenericData();
					
					for (k=0; k<names.length; k++) {
						gd.add(names[k], dp.getStringValue(names[k]));					
					}	
					
					managedSystem[i].addObject("proc", gd);
				}
				
				lines = sshm.sendCommand("lshwres -r procpool -m "+s[0], baseDir + s[0] + "_" + procPoolData);				
				br = new BufferedReader(new FileReader(baseDir + s[0] + "_" + procPoolData),1024*1024);
				
				names=null;
				
				for (j=0; j<lines; j++) {
					system = br.readLine();
					
					// Skip line if no data is returned
					if (system.startsWith("HSC") || system.startsWith("The managed system does not support"))
						continue;
					
					dp = new DataParser(system);
					names = dp.getNames();
					
					gd = new GenericData();
					
					for (k=0; k<names.length; k++) {
						gd.add(names[k], dp.getStringValue(names[k]));					
					}	
					
					managedSystem[i].addObject("procpool", gd);
				}
				
				
			} 
			catch (IOException ioe) {	
				System.out.println("Loader.getSystemData: IOException");
				System.out.println(ioe);
			}			
		}
		
		
		// Get MEM configuration of LPARs
		for (i=0; i<managedSystem.length; i++) {
			s = managedSystem[i].getVarValues("name");
			lines = sshm.sendCommand("lshwres -r mem --level lpar -m "+s[0], baseDir + s[0] + "_" + memLparData);
			
			try {			
				br = new BufferedReader(new FileReader(baseDir + s[0] + "_" + memLparData),1024*1024);
				
				names=null;
				
				for (j=0; j<lines; j++) {
					system = br.readLine();
					
					// Skip line if no data is returned
					if (system.startsWith("HSC"))
						continue;
					
					dp = new DataParser(system);
					names = dp.getNames();
					
					gd = new GenericData();
					
					for (k=0; k<names.length; k++) {
						gd.add(names[k], dp.getStringValue(names[k]));					
					}	
					
					managedSystem[i].addObject("mem", gd);
				}
				
				lines = sshm.sendCommand("lshwres -r mempool -m "+s[0], baseDir + s[0] + "_" + memPoolData);				
				br = new BufferedReader(new FileReader(baseDir + s[0] + "_" + memPoolData),1024*1024);
				
				names=null;
				
				for (j=0; j<lines; j++) {
					system = br.readLine();
					
					// Skip line if no data is returned
					if (system.startsWith("HSC"))
						continue;
					
					dp = new DataParser(system);
					names = dp.getNames();
					
					gd = new GenericData();
					
					for (k=0; k<names.length; k++) {
						gd.add(names[k], dp.getStringValue(names[k]));					
					}	
					
					managedSystem[i].addObject("mempool", gd);
				}
			} 
			catch (IOException ioe) {	
				System.out.println("Loader.getSystemData: IOException");
				System.out.println(ioe);
			}			
		}		
		
		
		

		
		
		names = managedSystem[0].getVarNames();
		for (i=1; i<managedSystem.length; i++)
			names = mergeList(names, managedSystem[i].getVarNames());
			
		
			
		WritableSheet sheet = workbook.createSheet("Systems", 0);
		
		// Create a cell format for Arial 10 point font
		WritableFont arial10font = new WritableFont(WritableFont.ARIAL, 10);
		WritableFont arial10boldfont = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD, true);
		WritableCellFormat arial10format = new WritableCellFormat (arial10font);
		WritableCellFormat arial10boldformat = new WritableCellFormat (arial10boldfont);
		
		WritableCellFormat diagonal = new WritableCellFormat (arial10boldfont);
		
		
		Label label;
		String v;

		
		try {			
			
			label = new Label(0,0,"HMC "+hmc+" Managed Systems",arial10boldformat);
			sheet.addCell(label);
			
			
			for (i=0; i<names.length; i++) {
				label = new Label(0,2+i,names[i],arial10boldformat);
				sheet.addCell(label);
			}
			
			diagonal.setOrientation(Orientation.PLUS_45);
			for (i=0; i<managedSystem.length; i++) {
				label = new Label(1+i,1,managedSystem[i].getVarValues("name")[0],diagonal);
				sheet.addCell(label);
			}
			
			
			
			for (i=0; i<managedSystem.length; i++) {
				
				for (j=0; j<names.length; j++) {
					s = managedSystem[i].getVarValues(names[j]);
					if (s==null)
						continue;
					
					v=null;
					for (k=0; k<s.length; k++) {	
						if (v==null)
							v=s[k];
						else
							v=v+","+s[k];
					}
				
					label = new Label(1+i,2+j,v,arial10format);
					sheet.addCell(label);
					
				}
				
				
			}
			
		} 
		catch (RowsExceededException ree) {	}
		catch (WriteException we) { }

			
		
		
		String systemName;
		GenericData object[];
		int row;
		
		sheet = workbook.createSheet("Slots",1);
		row = 0;
		
		// Merge variable names
		names=null;
		for (i=0; i<managedSystem.length; i++) {		
			object = managedSystem[i].getObjects("slot");
			if (object==null)
				continue;			
			for (j=0; j<object.length; j++)  {
				s = object[j].getVarNames();
				if (names==null)
					names=s;
				else
					names = mergeList(names, s);
			}
		}
		
		
		
		
		row=0;
		for (i=0; i<managedSystem.length; i++) {
			systemName = managedSystem[i].getVarValues("name")[0];
			
			try {
				label = new Label(0,row,systemName+" Slots",arial10boldformat);
				sheet.addCell(label);
				row++;
				
				object = managedSystem[i].getObjects("slot");
				if (object==null) {
					row+=2;
					continue;
				}
				
				// Print variable names				
				for (j=0; j<names.length; j++) {
					label = new Label(j,row,names[j],arial10boldformat);
					sheet.addCell(label);					
				}
				row++;
				
				for (j=0; j<object.length; j++) {
					for (k=0; k<names.length; k++) {
						s = object[j].getVarValues(names[k]);
						if (s==null)
							continue;
						
						v=null;
						for (x=0; x<s.length; x++) {	
							if (v==null)
								v=s[x];
							else
								v=v+","+s[x];
						}
					
						label = new Label(k,row,v,arial10format);
						sheet.addCell(label);
					}
					row++;
				}
				
				row+=2;				
			}
			catch (RowsExceededException ree) {	}
			catch (WriteException we) { }			
		}
		
		
		
		sheet = workbook.createSheet("CPU",2);
		row = 0;
		
		// Merge variable names
		names=null;
		for (i=0; i<managedSystem.length; i++) {		
			object = managedSystem[i].getObjects("proc");
			if (object==null)
				continue;			
			for (j=0; j<object.length; j++)  {
				s = object[j].getVarNames();
				if (names==null)
					names=s;
				else
					names = mergeList(names, s);
			}
		}
			
		row=0;
		for (i=0; i<managedSystem.length; i++) {
			systemName = managedSystem[i].getVarValues("name")[0];
			
			try {
				label = new Label(0,row,systemName+" CPU",arial10boldformat);
				sheet.addCell(label);
				row++;
				
				object = managedSystem[i].getObjects("proc");
				if (object==null) {
					row+=2;
					continue;
				}
				
				// Print variable names				
				for (j=0; j<names.length; j++) {
					label = new Label(j,row,names[j],arial10boldformat);
					sheet.addCell(label);					
				}
				row++;
				
				for (j=0; j<object.length; j++) {
					for (k=0; k<names.length; k++) {
						s = object[j].getVarValues(names[k]);
						if (s==null)
							continue;
						
						v=null;
						for (x=0; x<s.length; x++) {	
							if (v==null)
								v=s[x];
							else
								v=v+","+s[x];
						}
					
						label = new Label(k,row,v,arial10format);
						sheet.addCell(label);
					}
					row++;
				}
				
				object = managedSystem[i].getObjects("procpool");
				if (object!=null) {
					row++;
					
					String names2[]=null;
					for (j=0; j<object.length; j++)  {
						s = object[j].getVarNames();
						if (names2==null)
							names2=s;
						else
							names2 = mergeList(names2, s);
					}
					
					for (j=0; j<names2.length; j++) {
						label = new Label(j,row,names2[j],arial10boldformat);
						sheet.addCell(label);					
					}
					row++;
					
					for (j=0; j<object.length; j++) {
						for (k=0; k<names2.length; k++) {
							s = object[j].getVarValues(names2[k]);
							if (s==null)
								continue;
							
							v=null;
							for (x=0; x<s.length; x++) {	
								if (v==null)
									v=s[x];
								else
									v=v+","+s[x];
							}
						
							label = new Label(k,row,v,arial10format);
							sheet.addCell(label);
						}
						row++;
					}					
				}
				
				row+=2;				
			}
			catch (RowsExceededException ree) {	}
			catch (WriteException we) { }
			
			
		}
		
		
		sheet = workbook.createSheet("Memory",3);
		row = 0;
		
		// Merge variable names
		names=null;
		for (i=0; i<managedSystem.length; i++) {		
			object = managedSystem[i].getObjects("mem");
			if (object==null)
				continue;			
			for (j=0; j<object.length; j++)  {
				s = object[j].getVarNames();
				if (names==null)
					names=s;
				else
					names = mergeList(names, s);
			}
		}
			
		row=0;
		for (i=0; i<managedSystem.length; i++) {
			systemName = managedSystem[i].getVarValues("name")[0];
			
			try {
				label = new Label(0,row,systemName+" MEM",arial10boldformat);
				sheet.addCell(label);
				row++;
				
				object = managedSystem[i].getObjects("mem");
				if (object==null) {
					row+=2;
					continue;
				}
				
				// Print variable names				
				for (j=0; j<names.length; j++) {
					label = new Label(j,row,names[j],arial10boldformat);
					sheet.addCell(label);					
				}
				row++;
				
				for (j=0; j<object.length; j++) {
					for (k=0; k<names.length; k++) {
						s = object[j].getVarValues(names[k]);
						if (s==null)
							continue;
						
						v=null;
						for (x=0; x<s.length; x++) {	
							if (v==null)
								v=s[x];
							else
								v=v+","+s[x];
						}
					
						label = new Label(k,row,v,arial10format);
						sheet.addCell(label);
					}
					row++;
				}
				
				object = managedSystem[i].getObjects("mempool");
				if (object!=null) {
					row++;
					
					String names2[]=null;
					for (j=0; j<object.length; j++)  {
						s = object[j].getVarNames();
						if (names2==null)
							names2=s;
						else
							names2 = mergeList(names2, s);
					}
					
					for (j=0; j<names2.length; j++) {
						label = new Label(j,row,names2[j],arial10boldformat);
						sheet.addCell(label);					
					}
					row++;
					
					for (j=0; j<object.length; j++) {
						for (k=0; k<names2.length; k++) {
							s = object[j].getVarValues(names2[k]);
							if (s==null)
								continue;
							
							v=null;
							for (x=0; x<s.length; x++) {	
								if (v==null)
									v=s[x];
								else
									v=v+","+s[x];
							}
						
							label = new Label(k,row,v,arial10format);
							sheet.addCell(label);
						}
						row++;
					}					
				}
				
				row+=2;				
			}
			catch (RowsExceededException ree) {	}
			catch (WriteException we) { }
			
			
		}		
		
		


		try {
			workbook.write();
			workbook.close();
		} 
		catch (IOException ioe) { }
		catch (WriteException we) { } 
	
		
		
		
	}
	
	*/
	
	
	private String getActiveProfileName(int ms, String lparName) {
		GenericData lpar[];
		int i;
		
		lpar = managedSystem[ms].getObjects(CONFIG_LPAR);
		for (i=0; i<lpar.length; i++)
			if (lpar[i].getVarValues("name")[0].equals(lparName))
				break;
		
		if (i==lpar.length)
			return null;
		
		return lpar[i].getVarValues("curr_profile")[0];		
	}
	
	private GenericData getProfileData(int ms, String lparName, String profile) {
		GenericData profiles[];
		int i;
		
		if (lparName==null || profile==null)
			return null;
		
		profiles = managedSystem[ms].getObjects(PROFILES);
		if (profiles==null)
			return null; 
		
		for (i=0; i<profiles.length; i++) {
			if ( profiles[i].getVarValues("lpar_name")[0].equals(lparName) &&
					profiles[i].getVarValues("name")[0].equals(profile))
				break;
		}
		
		if (i<profiles.length)
			return profiles[i];
		
		return null;
	}
	
	private String[][] getProfilesUsingAdapter(int ms, String drcindex) {
		GenericData profiles[];
		Vector<String> required = new Vector<String>();
		Vector<String> desired = new Vector<String>();
		String slots[];
		String name;
		
		int i,j;
		
		if (drcindex==null)
			return null;
		
		profiles = managedSystem[ms].getObjects(PROFILES);
		if (profiles==null)
			return null;
		
		for (i=0; i<profiles.length; i++) {
			slots=profiles[i].getVarValues("io_slots");
			for (j=0; slots!=null && j<slots.length; j++) {
				if (slots[j].startsWith(drcindex)) {
					name = profiles[i].getVarValues("lpar_name")[0]+"@"+profiles[i].getVarValues("name")[0];
					if (slots[j].endsWith("1"))
						required.add(name);
					else
						desired.add(name);
				}
				
			}
		}
		
		String result[][] = new String[2][];
		result[1] = new String[required.size()];
		result[0] = new String[desired.size()];
		
		for (i=0; i<required.size(); i++)
			result[1][i]=required.elementAt(i);
		for (i=0; i<desired.size(); i++)
			result[0][i]=desired.elementAt(i);
		
		return result;
	}
	
	
	
	private PrintWriter createFrame(String fileName, String frameName) {
		PrintWriter html=null;
			
		try {
			html = new PrintWriter(
					new FileOutputStream(
					new File(fileName)));
		}
		catch (IOException e) { return null; }
		
		html.println("<HTML>\n" +
							"<HEAD><TITLE>"+frameName+"</TITLE></HEAD>\n" +
							"<BODY bgcolor=#cccccc>\n" +
							"<H1>"+frameName+"</H1>\n" +
							"<TABLE>\n");			
		
		return html;
	}
	
	private void closeFrame(PrintWriter html) {
		html.println("</TABLE>\n</BODY>\n</HTML>");
	}
	
	private void addTableRow(PrintWriter html, String text[]) {
		html.println("<TR>\n");
		for (int i=0; i<text.length; i++)
			if (text[i]==null)
				html.println("\t<TD></TD>\n");
			else
				html.println("\t<TD>"+text[i]+"</TD>\n");
		html.println("</TR>\n");
	}
	
	
	private void sendCommand(SSHManager2 sshm, String cmd, String fileName, boolean compressed) {
		sendCommand(sshm, cmd, fileName, compressed, false);
	}
	
	private void sendCommand(SSHManager2 sshm, String cmd, String fileName) {
		sendCommand(sshm, cmd, fileName, false, false);
	}
	
	
	private void sendCommand(SSHManager2 sshm, String cmd, String fileName, boolean compressed, boolean progress) {
		
		BufferedReader br;
		int i;
		String line;
		
		for (i=0; i<NUM_RETRY; i++) {
			
			// Send command to HMC
			sshm.sendCommand(cmd, fileName, compressed, progress);
			
			// Check if HMC lock happened
			
			try {
				br = new BufferedReader(new FileReader(fileName),1024*1024);
			} catch (FileNotFoundException fnfe) {
				break;
			}
			
			try {
				line=br.readLine();			
				
				if (line!=null && line.indexOf("Service processor lock failed")>=0) {
					// A lock occurred !!! Log and retry!!!
					System.out.print("X");
					br.close();
					
					try {
						Thread.sleep(10000);
					} catch (InterruptedException ie) {};
					
				} else {
					// No lock or empty file so stop
					br.close();
					break;
				}
			} catch (IOException ioe) {
				break;
			}				
		}
	}


}
