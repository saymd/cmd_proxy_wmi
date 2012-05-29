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
import java.util.logging.Level;

import net.saymd.addon.WindowsMachine;

import org.jinterop.dcom.common.JIException;
import org.jinterop.dcom.common.JISystem;
import org.jinterop.dcom.core.JIComServer;
import org.jinterop.dcom.core.JIProgId;
import org.jinterop.dcom.core.JISession;

public class AbstractComObject {
	protected JISession session;
	protected JIComServer comServer;
	private WindowsMachine machine;
	private JIProgId progId;

	/**
	 * initialize COM session.
	 * in this method,the COM server dose not been connected 
	 * @param machine target machine to be connected
	 * @param progId target COM to be connected
	 */
	public AbstractComObject(WindowsMachine machine, JIProgId progId) {
		this.machine = machine;
		JISystem.getLogger().setLevel(Level.INFO);
        
		session = JISession.createSession( machine.domain,
										   machine.username,
										   machine.password);
		session.useSessionSecurity(true);
		session.setGlobalSocketTimeout(5000);
		
		this.progId = progId;
	}

	/**
	 * connect to the COM server
	 * @throws UnknownHostException
	 * @throws JIException
	 */
	public void connect() throws UnknownHostException, JIException {
		progId.setAutoRegistration(true);
		comServer = new JIComServer(progId, machine.host, session);
	}

	/**
	 * disconnect session from the DCOM server
	 * @throws JIException
	 */
	public void disconnect() throws JIException {
		if (session != null) {
			JISession.destroySession(session);
		}
	}
}
