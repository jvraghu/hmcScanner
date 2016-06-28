package hmcScanner;

public class DiskData {
	
	private static byte STEP = 30;

	
	private int num_uuid = 0;	
	
	private int		uuid[] = null;
	private int		size[] = null;
	private int		u_vios[][] = null;
	private int		u_name[][] = null;
	private int		u_free[][] = null;
	
	
	private static byte		UUID 		= 0;
	private static byte		VIOSNAME 	= 1;
	private static byte		HDISKNAME 	= 2;
	private static byte		NUM_NAMES 	= 3;
	
	private int counter[] = new int[NUM_NAMES];
	private String name[][] = new String[NUM_NAMES][];
	private int index[][] = new int[NUM_NAMES][];
	
	
	public void addSize(String ms, String vios, String hdisk, int mb) {
		int hdisk_index = getName(hdisk, HDISKNAME, 0, counter[HDISKNAME]-1);
		int vios_index = getName(ms+"@"+vios, VIOSNAME, 0, counter[VIOSNAME]-1);
		int i,j;
		
		for (i=0; i<num_uuid; i++)
			for (j=0; j<u_vios[i].length; j++)
				if (u_vios[i][j]==vios_index && u_name[i][j]==hdisk_index) {
					size[i]=mb;
					return;
				}
	}
	
	public void addFree(String ms, String vios, String hdisk) {
		int hdisk_index = getName(hdisk, HDISKNAME, 0, counter[HDISKNAME]-1);
		int vios_index = getName(ms+"@"+vios, VIOSNAME, 0, counter[VIOSNAME]-1);
		int i,j;
		
		for (i=0; i<num_uuid; i++)
			for (j=0; j<u_vios[i].length; j++)
				if (u_vios[i][j]==vios_index && u_name[i][j]==hdisk_index) {
					u_free[i][j]=1;
					return;
				}
	}
	
	
	public int getNumUUID() {
		return num_uuid;
	}
	
	public String getUUIDname(int id) {
		return name[UUID][index[UUID][id]];
	}
	
	public int getSize(int id) {
		return size[index[UUID][id]];
	}
	
	
	public String[] getViosNames() {
		String result[] = new String[counter[VIOSNAME]];
		for (int i=0; i<counter[VIOSNAME]; i++)
			result[i] = name[VIOSNAME][index[VIOSNAME][i]];
		return result;
	}
	
	public String[] getHdiskOnViosNames(int id) {
		String result[] = new String[counter[VIOSNAME]];
		
		int i,j;
		
		for (i=0; i<counter[VIOSNAME]; i++) {
			for (j=0; j<u_vios[index[UUID][id]].length; j++)
				if (index[VIOSNAME][i]==u_vios[index[UUID][id]][j]) {
					result[i] = name[HDISKNAME][u_name[index[UUID][id]][j]];
					break;
				}
		}
		
		return result;
	}
	
	public boolean[] getFreeOnViosNames(int id) {
		boolean result[] = new boolean[counter[VIOSNAME]];
		
		int i,j;
		
		for (i=0; i<counter[VIOSNAME]; i++) {
			for (j=0; j<u_vios[index[UUID][id]].length; j++)
				if (index[VIOSNAME][i]==u_vios[index[UUID][id]][j]) {
					result[i] = (u_free[index[UUID][id]][j]==1?true:false);
					break;
				} 
		}
		
		return result;
	}
	
	
	public void add_uuid(String new_uuid, String new_hdisk, String ms, String vios) {
		int uuid_index = getName(new_uuid, UUID, 0, counter[UUID]-1);
		int hdisk_index = getName(new_hdisk, HDISKNAME, 0, counter[HDISKNAME]-1);
		int vios_index = getName(ms+"@"+vios, VIOSNAME, 0, counter[VIOSNAME]-1);
		
		if (num_uuid==0) {
			uuid = new int[STEP];
			size = new int[STEP];
			u_vios = new int[STEP][];
			u_name = new int[STEP][];
			u_free = new int[STEP][];			
		}
		
		int i;
		
		for (i=0; i<num_uuid && uuid[i]!=uuid_index; i++)
			;
		
		if (i==num_uuid) {
			// New entry!
			if (num_uuid==uuid.length) {
				int tmp_uuid[] = new int[uuid.length+STEP];
				int tmp_size[] = new int[uuid.length+STEP];
				int tmp_u_vios[][] = new int[uuid.length+STEP][];
				int tmp_u_name[][] = new int[uuid.length+STEP][];
				int tmp_u_free[][] = new int[uuid.length+STEP][];
				
				for (int j=0; j<uuid.length; j++) {
					tmp_uuid[j]=uuid[j];
					tmp_size[j]=size[j];
					tmp_u_vios[j]=u_vios[j];
					tmp_u_name[j]=u_name[j];
					tmp_u_free[j]=u_free[j];					
				}
				
				uuid = tmp_uuid;
				size = tmp_size;
				u_vios = tmp_u_vios;
				u_name = tmp_u_name;
				u_free = tmp_u_free;
			}
			
			uuid[num_uuid] = uuid_index;
			u_name[num_uuid] = addToList(hdisk_index,u_name[num_uuid]);
			u_vios[num_uuid] = addToList(vios_index,u_vios[num_uuid]);
			u_free[num_uuid] = addToList(0,u_free[num_uuid]);
			
			num_uuid++;
		} else {
			u_name[i] = addToList(hdisk_index,u_name[i]);
			u_vios[i] = addToList(vios_index,u_vios[i]);
			u_free[i] = addToList(0,u_free[i]);
		}
		
	}
	
	
	private int[] addToList(int n, int list[]) {
		int new_list[];
		
		if (list==null) {
			new_list = new int[1];
			new_list[0]=n;
			return new_list;
		}
		
		new_list = new int[list.length+1];
		for (int i=0; i<list.length; i++)
			new_list[i]=list[i];
		new_list[list.length]=n;
		return new_list;
	}
	
	
	/*
	 *  Search name in ordered list.
	 *  Returns:	position of name in list
	 */
	
	private int getName(String target, byte type, int from, int to) {
		if (counter[type]==0) {
			name[type]=new String[STEP];
			name[type][0]=target;
			index[type]=new int[STEP];
			index[type][0]=0;
			counter[type]=1;
			return 0;
		}
		
		
		int res = name[type][index[type][(from+to)/2]].compareTo(target);
		if (res==0)
			return index[type][(from+to)/2];
		
		if (from==to) {
			// New entry: add it to list and update index
			int i;
			if (counter[type]==name[type].length) {
				String new_name[] = new String[name[type].length+STEP];
				for (i=0; i<counter[type]; i++)
					new_name[i]=name[type][i];
				name[type]=new_name;
				int new_index[] = new int[name[type].length+STEP];
				for (i=0; i<counter[type]; i++)
					new_index[i]=index[type][i];
				index[type]=new_index;
			}
			name[type][counter[type]]=target;
			if (res<0) {
				for (i=counter[type]; i>from; i--)
					index[type][i]=index[type][i-1];
				index[type][from]=counter[type];
			} else {
				for (i=counter[type]; i>from+1; i--)
					index[type][i]=index[type][i-1];
				index[type][from+1]=counter[type];
			}
			counter[type]++;
			return counter[type]-1;
		}
		
		
		if (res<0)
			return getName(target, type, from, (from+to)/2);
		
		if (to==from+1)
			return getName(target, type, to, to);
		
		return getName(target, type, (from+to)/2, to);
		
	}
	

}
