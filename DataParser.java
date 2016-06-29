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

import java.util.Vector; 

public class DataParser {
	
	private String 	line = null;
	
	private Vector<String> name = new Vector<String>();
	private Vector<String[]> value = new Vector<String[]>();
	
	
	public DataParser(String line) {
		int		n;
		
		this.line = line;
		
		n=0;

		
		while (true) {
			n = parseVariableNew(n);
			if (n<0)
				break;
			
			if (n>=line.length()-1)
				break;
			
			if (line.charAt(n+1)==',')
				n+=2;
			else {
				System.out.println("DataParser: missing comma at position "+n+1);
				System.out.println(line.substring(0, n+1)+">>>"+line.charAt(n+1)+"<<<"+line.substring(n+2));
				System.out.println("");
				break;
			}
		}
		
		
	}
	
	/*
	 * Parse variable=value and return latest char pos
	 */
	private int parseVariableOld(int currPos) {
		
		boolean			multipleValues = false;
		int 			equals;
		
		String			variableName = null;
		String			variableValues[] = null;
		
		String			s;
		String			sa[];
		
		int 			i;
		int				n;
		
		
		if (currPos >= line.length()) {
			System.out.println("DataParser.parseVariable: currPos="+currPos+"; line.lenght()="+line.length());
			System.out.println(line.substring(0, currPos)+">>>"+line.charAt(currPos)+"<<<"+line.substring(currPos+1));
			System.out.println(line);
			System.out.println("");
			return -1;
		}

		
		if (line.charAt(currPos) == '"') {
			multipleValues = true;		
			currPos++;
		}
		
		// Detect variable name
		equals = line.indexOf('=', currPos);
		if (equals == -1) {
			// There is no variable. Skip this line.
			return -1;
		}
			
		variableName = line.substring(currPos, equals);
		currPos = equals+1;
		
		
		// bux fix for "description=bla, bla, bla," while it is a single variable value
		if (variableName.equals("description") && multipleValues) {
			n =  line.indexOf('"', currPos);
			variableValues = new String[1];
			variableValues[0] = line.substring(currPos, n);
			name.add(variableName);
			value.add(variableValues);
			return n;
		}
		
		
		// Check if empty variable
		if (currPos == line.length() || line.charAt(currPos)==',') {
			name.add(variableName);
			variableValues = new String[1];
			variableValues[0]="";
			value.add(variableValues);
			
			return currPos-1;
		}
		
		
		variableValues = new String[1];
		n=parseValueOld(currPos);
		variableValues[0] = line.substring(currPos, n+1);
		currPos=n;
		
		
		if (multipleValues) {
			currPos++;
			while (line.charAt(currPos) != '"') {
				
				// A comma must be present since it is a list
				if (line.charAt(currPos) != ',') {
					System.out.println("DataParser.parseVariable: Expected comma in pos="+currPos);
					System.out.println(line.substring(0, currPos)+">>>"+line.charAt(currPos)+"<<<"+line.substring(currPos+1));
					System.out.println("");
					return -1;
				} else
					currPos++;
				
				n=parseValueOld(currPos);
				s = line.substring(currPos, n+1);
				currPos=n+1;
				
				sa = variableValues;
				variableValues = new String[sa.length+1];
				for (i=0; i<sa.length; i++)
					variableValues[i]=sa[i];
				variableValues[sa.length]=s;
			}
		}
		
		
		name.add(variableName);
		value.add(variableValues);
		
		return currPos;
				
	}
	
	
	/*
	 * Parse variable=value and return latest char pos
	 */
	private int parseVariableNew(int currPos) {
		
		boolean			multipleValues = false;
		int 			equals;
		
		String			variableName = null;
		String			variableValues[] = null;
		
		String			s;
		String			sa[];
		
		int 			i;
		int				n;
		
		
		if (currPos >= line.length()) {
			System.out.println("DataParser.parseVariable: currPos="+currPos+"; line.lenght()="+line.length());
			System.out.println(line.substring(0, currPos)+">>>"+line.charAt(currPos)+"<<<"+line.substring(currPos+1));
			System.out.println(line);
			System.out.println("");
			return -1;
		}

		
		if (line.charAt(currPos) == '"') {
			multipleValues = true;		
			currPos++;
		}
		
		// Detect variable name
		equals = line.indexOf('=', currPos);
		if (equals == -1) {
			// There is no variable. Skip this line.
			System.out.println("DataParser.parseVariable: currPos="+currPos+"; line.lenght()="+line.length());
			System.out.println(line.substring(0, currPos)+">>>"+line.charAt(currPos)+"<<<"+line.substring(currPos+1));
			System.out.println(line);
			System.out.println("A variable name was expected. Skipping line.");
			System.out.println("");
			return -1;
		}
			
		variableName = line.substring(currPos, equals);
		currPos = equals+1;
		
		
		// bux fix for "description=bla, bla, bla," while it is a single variable value
		// same for "entry"
		if ( ( variableName.equals("description") || variableName.equals("entry") ) && multipleValues) {
			
			// Find a matching " but restart search if another " pair starts:    "description=bla bla""bla bla"
			n =  line.indexOf('"', currPos);
			while (n>0 && n<line.length()-1 && line.charAt(n+1)=='\"')
				n =  line.indexOf('"', n+2);
			if (n<0) {
				// No matching ", skip this line
				System.out.println("DataParser.parseVariable: currPos="+currPos+"; line.lenght()="+line.length());
				System.out.println(line.substring(0, currPos)+">>>"+line.charAt(currPos)+"<<<"+line.substring(currPos+1));
				System.out.println(line);
				System.out.println("Not matching \". Skipping line");
				System.out.println("");
				return -1;			
			}
			
			variableValues = new String[1];
			variableValues[0] = line.substring(currPos, n);
			name.add(variableName);
			value.add(variableValues);
			return n;
		}
		
		
		// Check if empty variable
		if (currPos == line.length() || line.charAt(currPos)==',') {
			name.add(variableName);
			variableValues = new String[1];
			variableValues[0]="";
			value.add(variableValues);
			
			return currPos-1;
		}
		
		
		variableValues = new String[1];
		n=parseValueNew(currPos);
		if (line.charAt(currPos)=='"')
			variableValues[0] = line.substring(currPos+2,n-1);
		else
			variableValues[0] = line.substring(currPos, n+1);
		currPos=n;
		
		
		if (multipleValues) {
			currPos++;
			while (line.charAt(currPos) != '"') {
				
				// A comma must be present since it is a list
				if (line.charAt(currPos) != ',') {
					System.out.println("DataParser.parseVariable: Expected comma in pos="+currPos);
					System.out.println(line.substring(0, currPos)+">>>"+line.charAt(currPos)+"<<<"+line.substring(currPos+1));
					System.out.println("");
					return -1;
				} else
					currPos++;
				
				n=parseValueNew(currPos);
				if (line.charAt(currPos)=='"')
					s = line.substring(currPos+2,n-1);
				else
					s = line.substring(currPos, n+1);
				//s = line.substring(currPos, n+1);
				currPos=n+1;
				
				sa = variableValues;
				variableValues = new String[sa.length+1];
				for (i=0; i<sa.length; i++)
					variableValues[i]=sa[i];
				variableValues[sa.length]=s;
			}
		}
		
		
		name.add(variableName);
		value.add(variableValues);
		
		return currPos;
				
	}
	
	
	
	
	/*
	 * Return the position of latest value char
	 */
	private int parseValueOld(int currPos) {
		int		n;
				
		if (line.charAt(currPos) == '"') {
			
			// This is a list of values
			n = currPos+1;
			while (true) {
				n = parseValueOld(n);
				if (n==line.length()-1) {
					System.out.println("DataParser.parseValue: Truncated line! I was scanning a list of values...");
					System.out.println(line);
					System.out.println("");
					return -1;
				}
				if (line.charAt(n+1) == ',')	n++;
				if (line.charAt(n+1) == '"')	break;				
			};
			return n+1;			
			
		} else {
			
			// This is a single value
			int comma, doubleQuotes, slash;
			
			comma = line.indexOf(',', currPos);
			doubleQuotes = line.indexOf('"', currPos);
			slash = line.indexOf('/', currPos);
			
			if (slash>0 
					&& ( doubleQuotes<0 || slash<doubleQuotes )
					&& ( comma <0 || slash<comma) ) {
				// Variable is made by multiple components divided by slash that may contain lists
				if (slash==line.length()-1)
					return slash;
				if (doubleQuotes==slash+1 || comma==slash+1)
					return slash;
				n = parseValueOld(slash+1);
				return n;				
			}
			
			if (comma>0 
					&& ( doubleQuotes<0 || comma<doubleQuotes) ) {
				// Comma identifies the end of the variable
				return comma-1;
			}
			
			if (doubleQuotes>0) {
				// Double quotes identifies the end of the variable
				return doubleQuotes-1;
			}
			
			// The variable ends with the line
			return line.length()-1;
			
		}
		
	}
	
	
	
	/*
	 * Detect variable value.
	 * currPos = beginning of value
	 * Returns: the position of the last character of value
	 */
	private int parseValueNew(int currPos) {
		int		n;
		
		// If two double quotes, variable end at the following couple of double quotes
		if (line.charAt(currPos) == '"' &&
				currPos+1 < line.length() &&
				line.charAt(currPos+1) == '"' ) {
			
			currPos = currPos+2;
			n = line.indexOf('"',currPos);
			if (n<0 || n+1 >= line.length() || line.charAt(n+1) != '"') {
				System.out.println("DataParser.parseValueNew: error parsing two double quotes!");
				System.out.println(line);
				System.out.println("index of first double quote = "+ (currPos-2) );
				return -1;				
			}
			return n+1;
		}
		
		// variable value ends when comma, double quote or EOL is found
		int comma = line.indexOf(',', currPos);
		int doubleQuotes = line.indexOf('"', currPos);
		
		// EOL
		if (comma<0 && doubleQuotes<0)
			return line.length()-1;
		
		if (comma<0)
			return doubleQuotes-1;
		
		if (doubleQuotes<0)
			return comma-1;
		
		if (comma<doubleQuotes)
			return comma-1;
		else
			return doubleQuotes-1;
	}
	
	
	
	
	public String[] getNames() {
		String s[]=new String[1];
		return name.toArray(s);
	}
	
	public String[] getStringValue(String key) {
		for (int i=0; i<name.size(); i++)
			if (key.equals(name.elementAt(i)))
				return value.elementAt(i);
		return null;
	}

}
