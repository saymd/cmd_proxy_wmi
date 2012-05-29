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
package net.saymd.addon.wmi;

import java.lang.reflect.Method;

import net.saymd.addon.AddonService;
import net.saymd.addon.WindowsMachine;
import net.saymd.addon.wmi.com.AbstractComObject;
import android.app.Activity;
import android.content.Intent;
import android.os.Bundle;

public class WmiProxyService extends AddonService {
	public WmiProxyService(String name) {
		super(name);
	}

	public WmiProxyService() {
		super(WmiProxyService.class.getSimpleName());
	}

	@Override
	protected void onPreExecute(Intent intent, Bundle param) {
		// put pod to each intent
		param.putSerializable( WmiProxyService.class.getSimpleName() + ".machine",
							  new WindowsMachine(intent));

		// overwrite parameters
		intent.putExtra("param", param);
		
		// add COM package prefix
		intent.putExtra("class", "com." + intent.getStringExtra("class") );
		
		super.onPreExecute(intent, param);
	}

	@Override
	protected void dispatchMethod(Class<?> clazz, Method requestedMethod,
			Bundle param) throws Exception {
		AbstractComObject comObject;

		comObject = (AbstractComObject) clazz
				.getConstructor( WindowsMachine.class)
				.newInstance( (WindowsMachine) param.getSerializable(
									WmiProxyService.class.getSimpleName() +".machine") 
									);

		comObject.connect();
		
		if (requestedMethod.getParameterTypes().length == 0) {
			requestedMethod.invoke(comObject, (Object[]) null);
		} else {
			requestedMethod.invoke(comObject, new Object[]{param});
		}

		comObject.disconnect();
		
		onPostExecute(param, Activity.RESULT_OK, null);
	}
}