/* WMI Proxy Service                                                                                                                                                      
 * Copyright (C) 2012  saymd                                                                                                                       
 *                                                                                                                                                                                       
 * This application is free software; you can redistribute it and/or                                                                                                                         
 * modify it under the terms of the GNU Lesser General Public                                                                                                                            
 * License as published by the Free Software Foundation; either                                                                                                                          
 * version 3.0 of the License, or (at your option) any later version.                                                                                                                    
 *                                                                                                                                                                                       
 * This application is distributed in the hope that it will be useful,                                                                                                                       
 * but WITHOUT ANY WARRANTY; without even the implied warranty of                                                                                                                        
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU                                                                                                                     
 * Lesser General Public License for more details.                                                                                                                                       
 *                                                                                                                                                                                       
 * You should have received a copy of the GNU Lesser General Public                                                                                                                      
 * License along with this application; if not, write to the Free Software                                                                                                                   
 * Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA                                                                                                             
 */
package net.saymd.addon.wmi.com;

import java.net.UnknownHostException;

import net.saymd.addon.WindowsMachine;

import org.jinterop.dcom.common.JIException;
import org.jinterop.dcom.core.IJIComObject;
import org.jinterop.dcom.core.JIProgId;
import org.jinterop.dcom.core.JIString;
import org.jinterop.dcom.impls.JIObjectFactory;
import org.jinterop.dcom.impls.automation.IJIDispatch;

import android.os.Bundle;
import android.util.Log;

public class WScriptShell extends AbstractComObject {
	private IJIDispatch shell;

	/**
	 * initialize WScript.ShellApplication
	 * @param machine target machine to be connected
	 */
	public WScriptShell(WindowsMachine machine) {
		super(machine, JIProgId.valueOf("WScript.Shell"));
	}
	
	@Override
	public void connect() throws UnknownHostException, JIException {
		super.connect();

		IJIComObject comUnknown = (IJIComObject) comServer.createInstance();
		shell = (IJIDispatch) JIObjectFactory
				.narrowObject((IJIComObject) comUnknown
						.queryInterface(IJIDispatch.IID));
	}

	/**
	 * execute command
	 * @param param command parameters. this object must have "command" key that binded String value
	 * @throws JIException
	 */
	public void execute(Bundle param) throws JIException {
		if (param.containsKey("command")) {
			String command = param.getString("command");
			
			if ( command != null && command != "") {
				Log.d("srv", command);
				shell.callMethodA("Exec", new Object[] { new JIString(command) });
			}
		}
	}
	
	/**
	 * run command
	 * @param param command parameters. this object must have "command" key that binded String value
	 * @throws JIException
	 */
	public void run(Bundle param) throws JIException {
		if (param.containsKey("command")) {
			String command = param.getString("command");
			
			if ( command != null && command != "") {
				shell.callMethodA("Run", new Object[] { new JIString(command) });
			}
		}
	}
	
	/**
	 * send keys to active window
	 * @param param command parameters. this object must have "key" key that binded String value
	 * @throws JIException
	 */
	public void sendKeys(Bundle param) throws JIException {
		if (param.containsKey("key")) {
			String key = param.getString("key");
			
			if ( key != null && key != "") {
				shell.callMethodA("SendKeys", new Object[] { new JIString(key) });
			}
		}
	}
}
