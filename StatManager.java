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


public class StatManager {
	
	private static final byte		BLOCKSIZE = 10;
	private double					step = 1;
	
	private int 					virtualCounter[][] = null;		// array of BLOCKSIZE
	private double					maxValue = 0;
	private int						numValues = 0;
	private double					sumValues = 0;
	
	
	private int getSubBlockId(double n) {
		int id;
		int newV[][];
		int i;
		
		id = (int)Math.ceil(n / step) / BLOCKSIZE;
		
		if (virtualCounter==null) {
			virtualCounter = new int[id+1][];
			virtualCounter[id] = new int[BLOCKSIZE];
			return id;
		}
		
		if (id>=virtualCounter.length) {
			newV = new int[id+1][];
			for (i=0; i<virtualCounter.length; i++)
				newV[i] = virtualCounter[i];
			newV[id] = new int[BLOCKSIZE];
			virtualCounter = newV;
			return id;
		}
		
		if (virtualCounter[id]==null) {
			virtualCounter[id] = new int[BLOCKSIZE];
			return id;
		}
		
		return id;
	}
	
	
	
	
	public void setStep(double step) {
		this.step = step;
	}


	public void addNumber(double n) {
		
		if (n>maxValue)
			maxValue = n;
		numValues++;
		sumValues += n;
		
		int blockId 	= getSubBlockId(n);
		int index 		= (int)Math.ceil(n/step) - blockId*BLOCKSIZE;
		
		virtualCounter[blockId][index]++;		
	}
	
	public double getMax() { return maxValue; }
	public double getAvg() { return numValues==0 ? -1 : sumValues/numValues; }
	
	public double getLevel(double level) {
		if (numValues==0)
			return 0;
		
		int target = (int)Math.ceil(numValues * level);
		int counter = 0;
		int i,j;
		
		i=j=0; 
		for (i=0; i<virtualCounter.length; i++)
			for (j=0; virtualCounter[i]!=null && j<BLOCKSIZE; j++) {
				counter += virtualCounter[i][j];
				if (counter>=target) {
					double result = step * ((i*BLOCKSIZE + j)+1); 
					if (result>maxValue)
						return maxValue;
					return result;
				}
			}
			
		double result = step * ((i*BLOCKSIZE + j)+1); 
		if (result>maxValue)
			return maxValue;
		return result;
	}
}
