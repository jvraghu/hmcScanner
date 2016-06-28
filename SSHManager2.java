package hmcScanner;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.Console;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.PrintWriter;
import java.util.zip.GZIPOutputStream;

import com.jcraft.jsch.Channel;
import com.jcraft.jsch.ChannelExec;
import com.jcraft.jsch.JSch;
import com.jcraft.jsch.JSchException;
import com.jcraft.jsch.KeyPair;
import com.jcraft.jsch.ProxyHTTP;
import com.jcraft.jsch.ProxySOCKS4;
import com.jcraft.jsch.ProxySOCKS5;
import com.jcraft.jsch.Session;
import com.jcraft.jsch.UserInfo;

public class SSHManager2 {
	
	private JSch 		ssh = null;
	private Session		session = null;
	private ProxyHTTP	proxyHTTP = null;
	private ProxySOCKS4	proxySOCKS4 = null;
	private ProxySOCKS5	proxySOCKS5 = null;
	private byte		proxyType = NONE;
	
	private static final byte NONE	= 0;
	private static final byte HTTP	= 1;
	private static final byte SOCKS4	= 2;
	private static final byte SOCKS5	= 3;
	
	
	
	private static boolean DEBUG = false; 
	private MyLogger myLogger = null;
	private int timeout = 0;
	
	
	
	
	// Set timeout in SECONDS
	public void setTimeout(int timeout) {
		this.timeout = timeout*1000;
	}

	@SuppressWarnings("unchecked")
	public static class MyLogger implements com.jcraft.jsch.Logger {
		@SuppressWarnings("rawtypes")
		static java.util.Hashtable name=new java.util.Hashtable();
		 static{
			 name.put(new Integer(DEBUG), "DEBUG: ");
			 name.put(new Integer(INFO), "INFO: ");
			 name.put(new Integer(WARN), "WARN: ");
			 name.put(new Integer(ERROR), "ERROR: ");
			 name.put(new Integer(FATAL), "FATAL: ");
		 }
		 private int level=FATAL;
		 private PrintWriter writer=null;
		 
		 public MyLogger(int level, String file) {
			 this.level=level;
			 try {
				 writer = new PrintWriter(new BufferedWriter(new FileWriter(file)));
			 } catch (IOException ioe) {
				System.out.println("MyLogger: could not create writer on "+file);
			}
		 }
		 public boolean isEnabled(int level){
			 if (level>=this.level)
				 return true;
			 return false;
		 }
		 public void log(int level, String message){
			 if (writer!=null) {
				 writer.println(name.get(new Integer(level))+message);
				 writer.flush();
			 }
		 }
		 
	 }
	
	
	
	public static class MyUserInfo implements UserInfo {
		
		private String password = null;
		
		
		public void setPassword(String password) {
			this.password = password;
		}

		public String getPassphrase() {
			// TODO Auto-generated method stub
			return null;
		}

		public String getPassword() {
			// TODO Auto-generated method stub
			return password;
		}

		public boolean promptPassword(String message) {
			// TODO Auto-generated method stub
			return true;
		}

		public boolean promptPassphrase(String message) {
			// TODO Auto-generated method stub
			return true;
		}

		public boolean promptYesNo(String message) {
			// TODO Auto-generated method stub
			return true;
		}

		public void showMessage(String message) {
			// TODO Auto-generated method stub
			
		}
		
	}
	
	
	
	public SSHManager2(String logFile) {
		if (logFile!=null) {
			myLogger = new MyLogger(com.jcraft.jsch.Logger.DEBUG,logFile);
			JSch.setLogger(myLogger);
		}
		ssh = new JSch();
	}
	
	
	
	
	
	public SSHManager2() {
		ssh = new JSch();
		//Logger.getLogger("com.sshtools").setLevel(Level.WARNING);		
	}
	
	
	public static void main(String[] args) {
		SSHManager2 sshm = new SSHManager2("c:\\aaa.log");
		/*
		if (!sshm.connect("9.71.196.28", "hscroot", "abc123")) {
			System.out.println("Connection failed");
			System.exit(0);
		}
		*/
		
		//sshm.setProxyHTTP("proxy.emea.ibm.com", 8080, null, null);
		if (!sshm.connectKey("9.71.196.28", "hscroot", "c:\\id_rsa.passphrase")) {
			System.out.println("Connection failed");
			System.exit(0);
		}
		
		int num = sshm.sendCommand("lshmc -v", "c:\\lssyscfg.txt");
		
		System.out.println("Lette "+num+" righe");
	
		
	}
	
	
	public void setProxyHTTP(String proxy_host, int proxy_port, String proxy_user, String proxy_password){
		proxyHTTP = new ProxyHTTP(proxy_host, proxy_port);
		if (proxy_user!=null)
			proxyHTTP.setUserPasswd(proxy_user, proxy_password);
		proxyType=HTTP;
		if (myLogger!=null)
			myLogger.log(com.jcraft.jsch.Logger.DEBUG, "PROXY_HTTP host="+proxy_host+"; port="+proxy_port+"; user="+proxy_user);
	}
	
	public void setProxySOCKS4(String proxy_host, int proxy_port, String proxy_user, String proxy_password){
		proxySOCKS4 = new ProxySOCKS4(proxy_host, proxy_port);
		if (proxy_user!=null)
			proxySOCKS4.setUserPasswd(proxy_user, proxy_password);
		proxyType=SOCKS4;
		if (myLogger!=null)
			myLogger.log(com.jcraft.jsch.Logger.DEBUG, "PROXY_SOCKS4 host="+proxy_host+"; port="+proxy_port+"; user="+proxy_user);
	}
	
	public void setProxySOCKS5(String proxy_host, int proxy_port, String proxy_user, String proxy_password){
		proxySOCKS5 = new ProxySOCKS5(proxy_host, proxy_port);
		if (proxy_user!=null)
			proxySOCKS5.setUserPasswd(proxy_user, proxy_password);
		proxyType=SOCKS5;
		if (myLogger!=null)
			myLogger.log(com.jcraft.jsch.Logger.DEBUG, "PROXY_SOCKS5 host="+proxy_host+"; port="+proxy_port+"; user="+proxy_user);
	}
	
	
		
	
	public boolean connectKey(String host, String user, String keyFile) {	
		KeyPair	keypair;
		
		try {
			keypair = KeyPair.load(new JSch(), keyFile);
			if (keypair.isEncrypted()) {
				// Need to ask for passphrase
				System.out.println("The provided private key is encrypted.");
				Console c = System.console();
		        if (c == null) {
		            System.err.println("No console available. Passphrase will be in clear text!");
		            BufferedReader in = new BufferedReader(new InputStreamReader(System.in));
		            String  passPhrase;
		            try {
		            	passPhrase = in.readLine();
		            } catch (IOException ioe) {
		            	return false;
		            }
		            if (!keypair.decrypt(passPhrase)) {
			        	System.out.println("Invalid passphrase.");
			        	return false;
			        }
		            ssh.addIdentity(keyFile, passPhrase);
		        } else {    
			        char passphrase[] = c.readPassword("Enter your passphrase: ");
			        byte pfbyte[] = new byte[passphrase.length];
			        for (int i=0; i<passphrase.length; i++)
			        	pfbyte[i] = (byte)passphrase[i];
			        if (!keypair.decrypt(pfbyte)) {
			        	System.out.println("Invalid passphrase.");
			        	return false;
			        }
			        ssh.addIdentity(keyFile, pfbyte);
		        }
			} else			
				ssh.addIdentity(keyFile);
						
			session = ssh.getSession(user, host, 22);
			if (timeout != 0)
				System.out.println("Setting timeout to user defined " + timeout/1000 + "seconds");
			session.setTimeout(timeout);
			java.util.Properties config = new java.util.Properties();
			config.put("StrictHostKeyChecking", "no");
			session.setConfig(config);
			
			switch (proxyType) {
				case HTTP:		session.setProxy(proxyHTTP); break;
				case SOCKS4:	session.setProxy(proxySOCKS4); break;
				case SOCKS5:	session.setProxy(proxySOCKS5); break;
			}
			
			MyUserInfo ui = new MyUserInfo();  
	        session.setUserInfo(ui); 

			session.connect();
			  
			return session.isConnected();
			
		} catch (JSchException ioe) {
			System.out.println("Error in connecting to "+host+" with user "+user+ "and keyfile "+keyFile);
			return false;			
		}
	}
	
	
	public boolean connect(String host, String user, String password) {		
		try {
			session = ssh.getSession(user, host, 22);
			if (timeout != 0)
				System.out.println("Setting timeout to user defined " + timeout/1000 + " seconds");
			session.setTimeout(timeout);
			java.util.Properties config = new java.util.Properties();
			config.put("StrictHostKeyChecking", "no");
			session.setConfig(config);
			
			switch (proxyType) {
				case HTTP:		session.setProxy(proxyHTTP); break;
				case SOCKS4:	session.setProxy(proxySOCKS4); break;
				case SOCKS5:	session.setProxy(proxySOCKS5); break;
			}
			
			MyUserInfo ui = new MyUserInfo();  
	        ui.setPassword(password);  
	        session.setUserInfo(ui); 

			session.connect();
			  
			return session.isConnected();
			
		} catch (JSchException ioe) {
			System.out.println("Error in connecting to "+host+" with user "+user);
			System.out.println(ioe.getMessage());
			return false;			
		}		
	}
	
	public void disconnect() {
		session.disconnect();
	}
	
	public int sendCommand(String command, String file) {
		return sendCommand(command, file, false, false);
	}
	
	public int sendCommand(String command, String file, boolean compressed) {
		return sendCommand(command, file, compressed, false);
	}
	
	public int sendCommand(String command, String file, boolean compressed, boolean progress) {
		
		OutputStream 			outStream = null;
		InputStream 			inStream = null;
		String 					end = "END-OF-TRANSMISSION";
		PrintWriter 			writer = null;
		int						numLines=0;
		long					counter=0;
		final long				HASH = 1024*1024;
		
	
		try {
			
			if (!compressed)
				writer = new PrintWriter(new BufferedWriter(new FileWriter(file)));
			else
				writer = new PrintWriter(new GZIPOutputStream(new FileOutputStream(file)));


			if (myLogger!=null)
				myLogger.log(com.jcraft.jsch.Logger.DEBUG, "Sending command: " + command);
			
			command = "LANG=C "+command+"; echo "+end;
			
			Channel channel = session.openChannel("exec");
			((ChannelExec) channel).setCommand(command);
			channel.setInputStream(null);
			((ChannelExec) channel).setErrStream(System.err);
			inStream = channel.getInputStream();

			channel.connect();
			
			if (DEBUG) {
				System.out.println("Sending command: " + command);
				System.out.println("to file: " + file);
			}
			  
		
			byte buffer[] = new byte[20*1024]; 
			int read;
			boolean stop=false;
			String inputLine = "";
			String line;
			int    eol;
			  
			while( !stop && (read = inStream.read(buffer)) > 0) {
				inputLine += new String(buffer, 0 , read);		
				
				if (progress) {
					counter += read;
					while (counter>HASH) {
						System.out.print(".");
						counter -= HASH;
					}
				}
				 
				while ( (eol=inputLine.indexOf('\n'))>=0 ) {
					line = inputLine.substring(0,eol);
					if (line.endsWith(end)) {
						stop = true;
						eol = line.indexOf(end);
						if (eol>0) {
							writer.println(line.substring(0, eol));
							numLines++;	 
						}
						break;
					} 
					
					writer.println(line);
					numLines++;
					
					inputLine = inputLine.substring(eol+1);
	
				}
				  
				// String out = new String(buffer, 0, read);
				// System.out.print(out);			     
			     
			}
			  
			channel.disconnect();
			
			writer.close();
			
			if (progress) {
				if (counter>0)
					System.out.print(". ");
				else
					System.out.print(" ");
			}
			
			if (myLogger!=null)
				myLogger.log(com.jcraft.jsch.Logger.DEBUG, "Output written into " + file);
			
			return numLines;
		  
		} catch (IOException ioe) {
			if (DEBUG)
				System.out.println("IOException");
			if (writer!=null)
				writer.close();
			if (myLogger!=null)
				myLogger.log(com.jcraft.jsch.Logger.DEBUG, "IOException");
			return 0;
		}
		catch (JSchException ioe) {
			if (DEBUG)
				System.out.println("JSchException");
			if (writer!=null)
				writer.close();
			if (myLogger!=null)
				myLogger.log(com.jcraft.jsch.Logger.DEBUG, "JSchException");
			return 0;
		}
	}
}
