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

import java.lang.reflect.Method;

import org.jinterop.dcom.common.JIException;

import android.app.Activity;
import android.app.IntentService;
import android.content.Intent;
import android.os.Bundle;
import android.os.ResultReceiver;
import android.util.Log;

public abstract class AddonService extends IntentService {
	public AddonService(String name) {
		super(name);
	}

	@Override
	protected void onHandleIntent(Intent intent) {
		Bundle param = intent.getBundleExtra("param");
		param = param == null ? new Bundle() : param;
		
		Log.d("srv", "param :"+param.toString());
		
		onPreExecute(intent, param);
		onExcetute(intent, param);
	}

	/**
	 * execute previous-execution task
	 * @param intent
	 * @param param
	 */
	protected void onPreExecute(Intent intent, Bundle param) {
		ResultReceiver receiver = (ResultReceiver)intent.getParcelableExtra("receiver");
		
		// filter invalid callback
		if (receiver != null
				&& !(receiver instanceof ResultReceiver)) {
			intent.removeExtra("receiver");
		}else{
			param.putParcelable("receiver", receiver);
		}
	};

	/**
	 * execute requested task
	 * @param intent
	 * @param param
	 */
	private void onExcetute(Intent intent, Bundle param) {
		Class<?> clazz = null;
		
		//pickup requested class
		try {
			Log.d("srv", getPackageName() + "." + intent.getStringExtra("class"));
			
			clazz = Class.forName( getPackageName() + "." + intent.getStringExtra("class") );
			
			Log.d("srv", getPackageName() + "." + intent.getStringExtra("class") +" found");
			
		} catch (ClassNotFoundException e) {
			onError(param, e);
			return;
		}
		
		String requestedMethodName = intent.getStringExtra("method");
		if (requestedMethodName == null) {
			Log.d("srv", "method name not found");
			
			return;
		}

		//pickup requested method
		Method requestedMethod = null;
		for (Method method : clazz.getMethods()) {
			if (requestedMethodName.equals(method.getName())) {
				requestedMethod = method;
			}
		}

		if (requestedMethod != null) {
			try {
				dispatchMethod(clazz, requestedMethod, param);
			} catch (Exception e) {
				Log.d("srv", "method dispathing failed");
				onError(param, e);
			}
		}else{
			Log.d("srv", "method not found");
		}
	}

	/**
	 * notify error to callee by dispatching callback method
	 * @param param
	 * @param exception
	 */
	protected void onError(Bundle param, Throwable exception) {
		Bundle resultData = new Bundle();
		
		Throwable wrapper = exception;
		if( exception instanceof JIException ){
			wrapper = new Exception( exception.getMessage() );
			wrapper.setStackTrace( exception.getStackTrace() );
		}
		
		resultData.putSerializable("exception", wrapper );
		onPostExecute(param, Activity.RESULT_CANCELED, resultData);
	}

	/**
	 * dispatch callback method  
	 * @param param
	 * @param resultCode
	 * @param resultData
	 */
	protected void onPostExecute(Bundle param, int resultCode, Bundle resultData) {
		ResultReceiver callback = null;
		if (param.containsKey("receiver")) {
			callback = (ResultReceiver)param.get("receiver");
		}

		Log.d("srv", "result code: " + resultCode);
		Log.d("srv", "result data: " + (resultData == null ? "NULL" : resultData.toString() ) );
		
		if (callback != null) {
			if( resultData == null ){
				resultData = new Bundle();
			}
			
			resultData.putLong("reqId", param.getLong("reqId",0));
			callback.send(resultCode, resultData);
		}else{
			Log.d("srv", "callback not found");
		}
	}

	/**
	 * dispatch requested method
	 * 
	 * @param clazz target COM class object to be connected
	 * @param requestedMethod target method to be called
	 * @param param method parameters
	 * @throws Exception
	 */
	protected abstract void dispatchMethod(Class<?> clazz,
			Method requestedMethod, Bundle param) throws Exception;
}
