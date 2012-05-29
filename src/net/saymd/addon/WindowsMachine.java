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
package net.saymd.addon;

import java.io.Serializable;

import android.content.Intent;

public class WindowsMachine implements Serializable {
	private static final long serialVersionUID = -1981762024859215986L;
	public String host;
	public String username;
	public String password;
	public String domain;
	
	public WindowsMachine(){}
	
	public WindowsMachine(Intent intent){
		host = intent.getStringExtra("host");
		domain = intent.getStringExtra("domain");
		username = intent.getStringExtra("username");
		password = intent.getStringExtra("password");
		
		//replace null by blank string for failsafe
		domain = domain == null ? "" : domain;
	}
}
