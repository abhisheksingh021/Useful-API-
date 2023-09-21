/***************************************************
 *Copyright (C) 2000 MIND 
 *All rights reserved.
 *The information contained here in is confidential and 
 *proprietary to MIND and forms the part of the MIND 
 *CREATED BY		:	Ashish Mehta
 *Date			:	23/09/2014
 *Description	:	Property File Utility will return the reference for Propert file
 *
 *Project		:	RAS(REQUISITION APPROVAL SYSTEM )
 **********************************************************/

package com.mind.sers.util;

import java.io.FileInputStream;
import java.util.Properties;

public class PropertyFileLoader {
	public Properties getPropertyFile() {
		Properties properties = null;
		String strPropertyPath = "";
		FileInputStream propfile = null;
		try {

			strPropertyPath = new String(getClass().getResource(
					"../../../../../../Security/SERS.properties").getPath());
			
			System.out.println("strPropertyPath>>>>"+strPropertyPath);

			if (strPropertyPath != null) {
				strPropertyPath = strPropertyPath.replaceAll("%20", " ");

				propfile = new FileInputStream(strPropertyPath);

				if (propfile != null) {
					properties = new Properties();
					properties.load(propfile);
				}
			}

		} catch (Exception e) {
			System.out.println("Error while loading property file :: ");
			e.printStackTrace();
		} finally {
			try {
				propfile.close();
			} catch (Exception e) {
				System.out.println("Error while closing input stream :: " + e);
			}
		}

		return properties;
	}
}
