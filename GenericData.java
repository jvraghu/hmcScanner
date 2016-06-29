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

public class GenericData {
	
	private String		varName[] = null;
	private String		varValues[][] = null;
	
	private byte		dataId[] = null; 
	private GenericData	dataObjects[][] = null;
	
	
	
	public GenericData[] getObjects(byte id) {
		int i;
		
		if (dataId==null)
			return null;
		
		for (i=0; i<dataId.length; i++)
			if (id==dataId[i])
				return dataObjects[i];
		
		return null;
	}
	
	public void addObject(byte name, GenericData object) {
		
		int nameId;
		int i;
		
		
		if (dataId==null) {
			dataId = new byte[1];
			dataId[0]=name;
			dataObjects = new GenericData[1][];
			dataObjects[0] = new GenericData[1];
			dataObjects[0][0]=object;
			return;
		}
		
		for (nameId=0; nameId<dataId.length; nameId++)
			if (name==dataId[nameId])
				break;
		
		if (nameId==dataId.length) {
			// New entry!	
			byte s[];
			GenericData d[][];
			
			s=dataId;
			d=dataObjects;
			dataId = new byte[s.length+1];
			dataObjects = new GenericData[s.length+1][];
			dataObjects[s.length] = new GenericData[1];
					
			for (i=0; i<s.length; i++) {
				dataId[i]=s[i];
				dataObjects[i]=d[i];
			}
			
			dataId[nameId] = name;
			dataObjects[s.length][0]=object;
			return;
		}
		
		
		GenericData d[];
		
		d = dataObjects[nameId];
		dataObjects[nameId] = new GenericData[d.length+1];
		for (i=0; i<d.length; i++)
			dataObjects[nameId][i]=d[i];
		dataObjects[nameId][d.length] = object;		
	}
	
	
	
	/*
	 * Get the position where the entry is or should be placed
	 */
	private int getPosition(String name) {
		int begin,pos,end;
		int result;
		
		if (varName.length==1) {
			result = name.compareTo(varName[0]);
			if (result<0 || result==0)
				return 0;
			else 
				return 1;
		}
		
		
		
		begin = 0;
		end = varName.length-1;
		
		while (end-begin>1) {
			pos = (begin+end)/2;
			result = name.compareTo(varName[pos]);
			
			if (result==0)
				return pos;
			
			if (result<0)
				end=pos;
			else
				begin=pos;
		}
		
		if (begin==end)
			return begin;
		
		result = name.compareTo(varName[begin]);
		if (result<0 || result==0)
			return begin;
		
		result = name.compareTo(varName[end]);
		if (result<0 || result==0)
			return end;
		
		return end+1;		
	}
	
	
	
	public void add(String name, String values[]) {
		
		String		vn[];
		String		vv[][];
		int			i;
		int			pos;
		
		if (varName == null) {
			varName = new String[1];
			varName[0] = name;
			varValues = new String[1][];
			varValues[0]=values;
			return;
		} 
		
		// Search position in ordered list
		pos = getPosition(name);
		
		// If name exists, overwrite values
		if (pos<varName.length && varName[pos].equals(name)) {
			varValues[pos]=values;
			return;
		} 
		
		// Add new name in pos keeping an ordered list
		vn = varName;
		vv = varValues;
		
		varName = new String[vn.length+1];
		varValues = new String[vv.length+1][];
		for (i=0; i<pos; i++) {
			varName[i]=vn[i];
			varValues[i]=vv[i];			
		}
		varName[pos]=name;
		varValues[pos]=values;
		for (i=pos+1; i<varName.length; i++) {
			varName[i]=vn[i-1];
			varValues[i]=vv[i-1];
		}
	}

	public String[] getVarNames() {
		return varName;
	}

	public String[] getVarValues(String name) {
		int pos = getPosition(name);
		
		if (pos>=varValues.length)
			return null;
		
		if (varName[pos].equals(name))
			return varValues[pos];
		else
			return null;
	}

}
