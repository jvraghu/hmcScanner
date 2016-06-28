package hmcScanner;

public class StringChanger {
	
	private String[]	original 	= null;		// text label to be translated, in order of insertion
	private String		prefix 		= null;		// translation prefix
	private static byte SIZE		= 10;		// min allocation block
	private int			free;					// first flee entry in arrays
	
	private boolean		recursive = false;		// if true, use provided changers
	private String		regex;
	private StringChanger[]	changers;
	
	public StringChanger(String prefix) {
		this.prefix = prefix;
		
		original = new String[SIZE];
		
		for (int i=0; i<SIZE; i++)
			original[i]=null;
		
		free = 0;
	}
	
	public StringChanger(String regex, StringChanger changers[]) {
		recursive = true;
		this.regex = regex;
		this.changers = changers;		
	}
	
	// Locate string in array. Insert if not present. Return position
	private int locate (String s) {
		int i;
		
		if (s==null || s.equals(""))
			s = "void";
		
		for (i=0; i<free; i++)
			if (original[i].equals(s))
				break;
		
		if (i<free)
			return i;
		
		if (i<original.length) {
			original[i]=s;
			free++;
			return i;
		}
		
		String[] newOriginal = new String[original.length+SIZE];
		for (int j=0; j<original.length; j++)
			newOriginal[j]=original[j];
		original=newOriginal;
		original[i]=s;
		free++;
		return i;
	}
	
	public String translate(String from) {
		if (!recursive) {
			int pos = locate(from);
			return prefix + pos;
		}
		
		String	split[] = from.split(regex);
		for (int i=0; i<split.length; i++)
			split[i] = changers[i].translate(split[i]);
		
		String result = split[0];
		for (int i=1; i<split.length; i++)
			result = result + regex + split[i];
		return result;
	}

}
