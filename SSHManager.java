package hmcScanner;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintWriter;
import java.util.logging.FileHandler;
import java.util.logging.Handler;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.logging.SimpleFormatter;
import java.util.zip.GZIPOutputStream;

import com.sshtools.j2ssh.SshClient;
import com.sshtools.j2ssh.authentication.AuthenticationProtocolState;
import com.sshtools.j2ssh.authentication.PasswordAuthenticationClient;
import com.sshtools.j2ssh.authentication.PublicKeyAuthenticationClient;
import com.sshtools.j2ssh.session.SessionChannelClient;
import com.sshtools.j2ssh.transport.IgnoreHostKeyVerification;
import com.sshtools.j2ssh.transport.publickey.InvalidSshKeyException;
import com.sshtools.j2ssh.transport.publickey.SshPrivateKey;
import com.sshtools.j2ssh.transport.publickey.SshPrivateKeyFile;

public class SSHManager {
	
	private SshClient 		ssh = null;
	
	// Values: 
	//	AuthenticationProtocolState.FAILED
	//	AuthenticationProtocolState.PARTIAL
	//	AuthenticationProtocolState.COMPLETE
	
	private int				status = AuthenticationProtocolState.FAILED;
	
	private static boolean DEBUG = false; 
	
	
	
	public SSHManager(String logFile) {
		ssh = new SshClient();
		
		if (logFile != null) {
			// Setup a logfile
			Handler fh;
			try {
				fh = new FileHandler(logFile);
				fh.setFormatter(new SimpleFormatter());
				Logger.getLogger("com.sshtools").setUseParentHandlers(false);
				Logger.getLogger("com.sshtools").addHandler(fh);
				Logger.getLogger("com.sshtools").setLevel(Level.ALL);
				System.out.println("Started logging SSH connection to "+logFile);
				return;
			} catch (IOException e) {
				System.out.println("Logging SSH connection to "+logFile+" failed");
			}		
		}
		
		Logger.getLogger("com.sshtools").setLevel(Level.WARNING);	
	}
	
	
	
	
	
	public SSHManager() {
		ssh = new SshClient();
		Logger.getLogger("com.sshtools").setLevel(Level.WARNING);		
	}
	
	
	public static void main(String[] args) {
		SSHManager sshm = new SSHManager();
		if (!sshm.connect("127.0.0.1", "sysadmin", "passw0rd")) {
			System.out.println("Connection failed");
			System.exit(0);
		}
		
		int num = sshm.sendCommand("lsconfig -V", "c:\\lssyscfg.txt");
		
		System.out.println("Lette "+num+" righe");
	
		
	}
	
	
	public boolean connectKey(String host, String user, String keyFile) {		
		try {
			// Connect to host ignoring host key verification
			ssh.connect(host, new IgnoreHostKeyVerification());
			
			PublicKeyAuthenticationClient sshClient = new PublicKeyAuthenticationClient();

			SshPrivateKeyFile sshPrivKeyFile = SshPrivateKeyFile.parse(new File(keyFile));
			
			SshPrivateKey sshPrivKey = null;
			try {
				sshPrivKey = sshPrivKeyFile.toPrivateKey("");
			} catch (InvalidSshKeyException iske) {
				System.out.println("Wrong key or unsupported usage of passphrase.");
				return false;
			}
			sshClient.setKey(sshPrivKey);
			sshClient.setUsername(user);
			
			status = ssh.authenticate(sshClient);
			
			if (status == AuthenticationProtocolState.COMPLETE)
				return true;
			else
				return false;			
		} catch (IOException ioe) {
			System.out.println("Error in connecting to "+host+" with keyfile "+keyFile);
			System.out.println(ioe);
			return false;			
		}		
	}
	
	
	public boolean connect(String host, String user, String password) {		
		try {
			// Connect to host ignoring host key verification
			ssh.connect(host, new IgnoreHostKeyVerification());
			
			PasswordAuthenticationClient pwd = new PasswordAuthenticationClient();
			pwd.setUsername(user);
			pwd.setPassword(password);
			
			status = ssh.authenticate(pwd);
			
			if (status == AuthenticationProtocolState.COMPLETE)
				return true;
			else
				return false;			
		} catch (IOException ioe) {
			System.out.println("Error in connecting to "+host+" with user "+user);
			return false;			
		}		
	}
	
	public void disconnect() {
		ssh.disconnect();
	}
	
	public int sendCommand(String command, String file) {
		return sendCommand(command, file, false, false);
	}
	
	public int sendCommand(String command, String file, boolean compressed) {
		return sendCommand(command, file, compressed, false);
	}
	
	public int sendCommand(String command, String file, boolean compressed, boolean progress) {
		
		SessionChannelClient 	session = null;
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


		
			session = ssh.openSessionChannel();
			session.startShell();
			
			outStream = session.getOutputStream();
			end = "END-OF-TRANSMISSION";
			
			if (DEBUG) {
				System.out.println("Sending command: " + command);
				System.out.println("to file: " + file);
			}
			
			outStream.write( ("LANG=C "+command+"; echo "+end+"\n").getBytes());
			  
			/**
			  * Reading from the session InputStream
			  */
			inStream = session.getInputStream();
			
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
			  
			session.close();
			
			writer.close();
			
			if (progress) {
				if (counter>0)
					System.out.print(". ");
				else
					System.out.print(" ");
			}
			
			return numLines;
		  
		} catch (IOException ioe) {
			if (DEBUG)
				System.out.println("IOException");
			if (writer!=null)
				writer.close();
			return 0;
		}

		  
		  
		
	}
	
	
	

}
