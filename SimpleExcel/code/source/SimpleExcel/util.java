package SimpleExcel;

// -----( IS Java Code Template v1.2
// -----( CREATED: 2006-01-19 17:12:27 JST
// -----( ON-HOST: xiandros-c640

import com.wm.data.*;
import com.wm.util.Values;
import com.wm.app.b2b.server.Service;
import com.wm.app.b2b.server.ServiceException;
// --- <<IS-START-IMPORTS>> ---
import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import java.text.SimpleDateFormat;
import java.util.GregorianCalendar;
import com.wm.app.b2b.server.*;
// --- <<IS-END-IMPORTS>> ---

public final class util

{
	// ---( internal utility methods )---

	final static util _instance = new util();

	static util _newInstance() { return new util(); }

	static util _cast(Object o) { return (util)o; }

	// ---( server methods )---




	public static final void MSExcelWorkSheetToRecord (IData pipeline)
        throws ServiceException
	{
		// --- <<IS-START(MSExcelWorkSheetToRecord)>> ---
		// @subtype unknown
		// @sigtype java 3.5
		// [i] object:0:optional bindata
		// [i] field:0:optional xlsData
		// [i] field:0:optional encoding
		// [i] field:0:optional validate {"true","false"}
		// [i] field:0:optional returnErrors {"asArray","inResults","both"}
		// [o] record:0:required recordMSExcel
		// [o] - record:1:required recordMSExcel
		// [o] -- record:1:required row
		// [o] --- field:0:required C0
		// [o] --- field:0:required C1
		// [o] record:1:optional errors
		// [o] field:0:optional isValid {"true","false"}
		// pipeline
		IDataCursor pipelineCursor = pipeline.getCursor();
		
			String	file_name = null;
			BufferedInputStream	file_stream = null;
			ByteArrayInputStream    byte_array_input_stream = null;
			String	file_data = null;
		
			// data
			IData	data = IDataUtil.getIData( pipelineCursor, "filedata" );
			if ( data != null)
			{
				IDataCursor dataCursor = data.getCursor();
					file_name = IDataUtil.getString( dataCursor, "file_name" );
					file_stream = (BufferedInputStream)IDataUtil.get( dataCursor, "file_stream" );
					byte_array_input_stream = (ByteArrayInputStream)IDataUtil.get( dataCursor, "byte_array_stream" );
					file_data = IDataUtil.getString( dataCursor, "file_data" );
				dataCursor.destroy();
			}
		
		byte bin_array[] = (byte[])IDataUtil.get( pipelineCursor, "bindata" );
		String xlsData = IDataUtil.getString( pipelineCursor, "xlsData" );	
		String s_encoding = IDataUtil.getString( pipelineCursor, "encoding" );
		String s_validate = IDataUtil.getString( pipelineCursor, "validate" );
		String s_returnErrors = IDataUtil.getString( pipelineCursor, "returnErrors" );
		
		pipelineCursor.destroy();
		
		if (xlsData != null)
		    file_data = xlsData;
		
		// pipeline
		IDataCursor pipelineCursor_1 = pipeline.getCursor();
		IData[] work_sheet_list = null;
		IData[]	row_list = null;
		
		boolean is_valid = false;
		
		try
		{
			HSSFWorkbook wb = null;
			// Handle Inputs here
			if (file_name != null)
			{
			    if (file_name.length() > 0)
			    { 				
		    		wb = new HSSFWorkbook(new FileInputStream(file_name));
		            }
			}
			else if (file_stream != null)
			{		
				wb = new HSSFWorkbook(file_stream);
			}
			else if (file_data != null)
			{
				if (file_data.length() > 0)
				    wb = new HSSFWorkbook(new ByteArrayInputStream(file_data.getBytes()));
			}
			else if (bin_array != null)
			{		
				wb = new HSSFWorkbook(new ByteArrayInputStream(bin_array));
			}
			else if (byte_array_input_stream != null)
			{
				wb = new HSSFWorkbook(byte_array_input_stream);
			}
		
		    HSSFSheet sheet = null;
		    HSSFRow row = null;
		    HSSFCell cell = null;
		    String cl = "";
		    double icl = 0;
		    boolean bcl = false;
		
		    work_sheet_list = new IData[wb.getNumberOfSheets()];
		    IDataCursor	idc_sheet_node = null;
		    for (int ws = 0; ws < wb.getNumberOfSheets(); ws++)
		    {
			sheet = wb.getSheetAt(ws);
			work_sheet_list[ws] = IDataFactory.create();
			idc_sheet_node = work_sheet_list[ws].getCursor();
			//////////////////////////////////////////////////////////////
			// Read Excel data and create dynamic record based on fileds
		
			int row_cnt = sheet.getPhysicalNumberOfRows();
			row_list = new IData[row_cnt];
			
			for (int i = 0; i < row_cnt; i++)
			{
			     row = sheet.getRow(i);
			
			     IDataCursor idc_row_node = null;
			     row_list[i] = IDataFactory.create();
			     idc_row_node = row_list[i].getCursor();
		
			     if (row != null)
			     {
				int phys_cell_num = row.getPhysicalNumberOfCells();
				int last_cell_num = row.getLastCellNum();
				int total_cell = phys_cell_num;
				if (phys_cell_num < last_cell_num)
				    total_cell = last_cell_num;
		    	     	for (int j = 0; j < total_cell; j++)
		    	     	{
		        	     cell = row.getCell((short)j);
		        	     if (cell != null)
		        	     {
		
				    	switch (cell.getCellType()) 
				    	{
		              			case HSSFCell.CELL_TYPE_STRING:
		 	            			cl = cell.getStringCellValue();
							IDataUtil.put( idc_row_node, "C"+Integer.toString(j), cl );
		            			break;
		           	
		            			case HSSFCell.CELL_TYPE_NUMERIC:
		                			icl = cell.getNumericCellValue();
		                    			if (isCellDateFormatted(cell))
		                    			{
		                        	    	    // format in form of M/D/YY
		                        	    	    Calendar cal = Calendar.getInstance();
		                        	    	    cal.setTime(getJavaDate(icl,false));
		                        	    	    String pattern = getCellDateFormat(cell);
		                        	    	    SimpleDateFormat df = new SimpleDateFormat(pattern);
		                        	    	    String dateStr = df.format(cal.getTime());
						    
						    	    IDataUtil.put( idc_row_node, "C"+Integer.toString(j), dateStr );
		                    			}
		                    			else
		                		    	     IDataUtil.put( idc_row_node, "C"+Integer.toString(j), Double.toString(icl));
						break;
		
						case HSSFCell.CELL_TYPE_BOOLEAN:
							bcl = cell.getBooleanCellValue();
							if (bcl)
							    IDataUtil.put( idc_row_node, "C"+Integer.toString(j), "true" );
					    		else
					    		    IDataUtil.put( idc_row_node, "C"+Integer.toString(j), "false" );
					    	break;
		
					        case HSSFCell.CELL_TYPE_FORMULA:
						try 
						{	
							icl = cell.getNumericCellValue();
							// check if value is a NaN - NOT NUMBER
							if (!Double.isNaN(icl))
							    IDataUtil.put( idc_row_node, "C"+Integer.toString(j), Double.toString(icl));
							else
							{
							    cl = cell.getStringCellValue();
							    IDataUtil.put( idc_row_node, "C"+Integer.toString(j),cl);
							}
						} 
						catch(Exception fe) 
						{
					    		cl = cell.getCellFormula();
					    		IDataUtil.put( idc_row_node, "C"+Integer.toString(j), cl );
						}
					    	break;
		
		            			case HSSFCell.CELL_TYPE_BLANK:
			                		//icl = cell.getNumericCellValue();
							IDataUtil.put( idc_row_node, "C"+Integer.toString(j), "" );
						break;
		
		  			 	case HSSFCell.CELL_TYPE_ERROR:
					    		IDataUtil.put( idc_row_node, "C"+Integer.toString(j), "" );
						break;
		            		} // END SWITCH
				}
				else
				{
					IDataUtil.put( idc_row_node, "C"+Integer.toString(j), "" );
				}
		    	     } // End of J FOR
			   } // End of if row != null chek
			} //END for i
			// Setup worksheet record
			IDataUtil.put( idc_sheet_node, "row", row_list);
		
		    } // End of worsheet loop
		    is_valid = true;	
		}
		catch (Exception e)
		{
			is_valid = false;
			e.printStackTrace();
			throw new ServiceException(e.getMessage());
		}
		
		pipelineCursor_1.destroy();
		
		// pipeline
		IDataCursor pipelineCursor_2 = pipeline.getCursor();
		
		// recordMSExcel
		IData	recordMSExcel = IDataFactory.create();
		IDataCursor recordMSExcelCursor = recordMSExcel.getCursor();
		
		IDataUtil.put( recordMSExcelCursor, "recordMSExcel", work_sheet_list );
		recordMSExcelCursor.destroy();
		
		IDataUtil.put( pipelineCursor, "recordMSExcel", recordMSExcel );
		if (is_valid)
			IDataUtil.put( pipelineCursor, "isValid", "true" );
		else
			IDataUtil.put( pipelineCursor, "isValid", "true" );
		
		pipelineCursor.destroy();
		// --- <<IS-END>> ---

                
	}



	public static final void RecordToMSExcel (IData pipeline)
        throws ServiceException
	{
		// --- <<IS-START(RecordToMSExcel)>> ---
		// @subtype unknown
		// @sigtype java 3.5
		// [i] record:1:required in_doc
		// [i] field:0:optional file
		// [i] field:0:optional option {"file","bytes","stream"}
		// [o] object:0:optional bytes
		// [o] object:0:optional stream
		// [o] field:0:optional file_name
		// [o] field:0:required status
		// pipeline
		IDataCursor pipelineCursor = pipeline.getCursor();
			String	file = IDataUtil.getString( pipelineCursor, "file" );
			String opt = IDataUtil.getString( pipelineCursor, "option" );
			byte bin_array[] = (byte[])IDataUtil.get( pipelineCursor, "bindata" );
		
		String status = "false";
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet sheet = wb.createSheet("new sheet");
		HSSFRow row = null;
		HSSFCell cell = null;
		
		if (opt == null)
		    opt = "bytes";
		if (file != null)
		    opt = "file";
		try {
		
		    String key = "";
		    String val = "";
		    Object valObj = null;
		/*
		    // Create a cell.
		    row.createCell((short)0).setCellValue(1);
		    row.createCell((short)1).setCellValue(1.2);
		    row.createCell((short)2).setCellValue("This is a string");
		    row.createCell((short)3).setCellValue(true);
		*/
		    // in_doc
		    IData[] in_doc = IDataUtil.getIDataArray( pipelineCursor, "in_doc" );
		    if ( in_doc != null)
		    {
			// Handle all records - rows
			for ( int i = 0; i < in_doc.length; i++ )
			{
		    	     // Create a row and put some cells in it. Rows are 0 based.
		    	     row = sheet.createRow((short)i);
			     
			     int count = 0;
			     IDataCursor idc = in_doc[i].getCursor();
			     idc.first();
			     boolean more_data = true;
			     while (more_data)//(idc.hasMoreData())
			     {
				key = idc.getKey();
				val = (String)idc.getValue();
		    
				// Create a cell.
				//row.createCell((short)count).setCellValue(val);
				cell = row.createCell((short)count);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellValue(val);
				count++;
		
		    		// set status
		    		status = "true";
		
			        more_data = idc.next();
			     }
		             idc.destroy();
			}
		    }
		    pipelineCursor.destroy();
		
		if (opt.equals("file"))
		{
		    // Write the output to a file
		    FileOutputStream fileOut = new FileOutputStream(file);
		    wb.write(fileOut);
		    fileOut.close();
		    IDataCursor pipelineCursor2 = pipeline.getCursor();
		    IDataUtil.put( pipelineCursor2, "bytes", null );
		    IDataUtil.put( pipelineCursor2, "stream", null );	
		    pipelineCursor2.destroy();
		
		}
		else if (opt.equals("bytes"))
		{
		    ByteArrayOutputStream stream = new ByteArrayOutputStream();
		    wb.write(stream);
		    IDataCursor pipelineCursor2 = pipeline.getCursor();
		    IDataUtil.put( pipelineCursor2, "bytes", stream.toByteArray() );
		    pipelineCursor2.destroy();
		    stream.close();
		}
		else if (opt.equals("stream"))
		{
		    ByteArrayOutputStream stream = new ByteArrayOutputStream();
		    wb.write(stream);
		    IDataCursor pipelineCursor2 = pipeline.getCursor();
		    IDataUtil.put( pipelineCursor2, "stream", stream );	
		    pipelineCursor2.destroy();
		    stream.close();
		}
		
		} catch (Exception e) {
			//e.printStackTrace();
			pipelineCursor.destroy();
			throw new ServiceException(e.getMessage());	
		}
		
		// pipeline
		IDataCursor pipelineCursor1 = pipeline.getCursor();
		IDataUtil.put( pipelineCursor1, "file_name", file );
		IDataUtil.put( pipelineCursor1, "status", status );
		pipelineCursor1.destroy();
		// --- <<IS-END>> ---

                
	}



	public static final void get_dir (IData pipeline)
        throws ServiceException
	{
		// --- <<IS-START(get_dir)>> ---
		// @subtype unknown
		// @sigtype java 3.5
		// [o] field:0:required pkg_config
		// [o] field:1:required config_file_list
		IDataHashCursor idc = pipeline.getHashCursor();
		
			// Get input values
		   	idc.first();
			String key = idc.getKey();
			
			Values vl = ValuesEmulator.getValues(pipeline, key);
			String pkg = Service.getPackageName(vl);
			File fl = ServerAPI.getPackageConfigDir(pkg);
			String config_dir = fl.getPath();
		
			try
			{	
				// Get list of files in a give directory
		        File fname = new File(config_dir);
		        String[] file_list = fname.list();
				fname = null;
			   	idc.first();
				idc.insertAfter("config_file_list", file_list);
			}
			catch(Exception e)
			{
				throw new ServiceException(e.getMessage());
			}
		
			// Setup output message
			idc.first();
			idc.insertAfter("pkg_config",config_dir + File.separator);
		
			idc.destroy();
		// --- <<IS-END>> ---

                
	}



	public static final void query_cell (IData pipeline)
        throws ServiceException
	{
		// --- <<IS-START(query_cell)>> ---
		// @subtype unknown
		// @sigtype java 3.5
		// [i] object:0:optional bindata
		// [i] field:0:optional file_name
		// [i] field:0:optional worksheet_id
		// [i] field:0:required row_id
		// [i] field:0:required cell_id
		// [o] field:0:required cell_data
		
		// pipeline
		IDataCursor pipelineCursor = pipeline.getCursor();
			String	worksheet_id = IDataUtil.getString( pipelineCursor, "worksheet_id" );
			String	row_id = IDataUtil.getString( pipelineCursor, "row_id" );
			String	cell_id = IDataUtil.getString( pipelineCursor, "cell_id" );
			String file_name = IDataUtil.getString( pipelineCursor, "file_name" );
			byte bin_array[] = (byte[])IDataUtil.get( pipelineCursor, "bindata" );
		pipelineCursor.destroy();
		
		printdbg("## worksheet - " + worksheet_id + " row - "+row_id+" cell - "+cell_id);
		
		int worksheet = 0;
		int row = 0;
		short cell = 0;
		String value = "";
		// Set default worksheet id = 0
		if (worksheet_id != null)
		    worksheet = Integer.parseInt(worksheet_id);
		if (row_id != null)
		    row = Integer.parseInt(row_id);
		if (cell_id != null)
		    cell = Short.parseShort(cell_id);	
		
		try
		{
			HSSFWorkbook ms_wb = null;
			// Handle Inputs here
			if (file_name != null)
			{
			    if (file_name.length() > 0)
			    { 				
		    		ms_wb = new HSSFWorkbook(new FileInputStream(file_name));
		            }
			}
			else if (bin_array != null)
			{		
				ms_wb = new HSSFWorkbook(new ByteArrayInputStream(bin_array));
			}
			else throw new ServiceException("Input data is missing. bindata or file_name must be provided.");
		
		
			HSSFSheet sheet = null;
			HSSFRow ms_row = null;
			HSSFCell ms_cell = null;
			if (ms_wb != null)
			{			
				sheet = ms_wb.getSheetAt(worksheet);
				if (sheet != null)
				{
					printdbg("### queryExcelDocument: workbook-"+worksheet_id+" was created");
					ms_row = sheet.getRow(row);
					if (ms_row != null)
					{
						printdbg("### queryExcelDocument: row-"+row+" was created");
						ms_cell = ms_row.getCell(cell);
						if (ms_cell != null)
						{
							printdbg("### queryExcelDocument: cell-"+cell+" was created");
							int type = ms_cell.getCellType();
							double icl = 0;
							switch (type) 
							{
								case HSSFCell.CELL_TYPE_STRING:
									value = ms_cell.getStringCellValue();
									printdbg("### queryExcelDocument: cell type String -"+value);
								break;
		         	
								case HSSFCell.CELL_TYPE_NUMERIC:
									icl = ms_cell.getNumericCellValue();
									if (isCellDateFormatted(ms_cell))
									{
										// format in form of M/D/YY
										Calendar cal = Calendar.getInstance();
										cal.setTime(getJavaDate(icl,false));
										String pattern = "dd-MMM-yy";
										SimpleDateFormat df = new SimpleDateFormat(pattern);
										value = df.format(cal.getTime());
										printdbg("### queryExcelDocument: cell type Date -"+value);
									}
									else
									{
										Double dbl = new Double(icl);
										value = Long.toString(dbl.longValue());
										printdbg("### queryExcelDocument: cell type NUMERIC -"+value);
									}
								break;
									
								case HSSFCell.CELL_TYPE_BLANK:
									value = "";
									printdbg("### queryExcelDocument: cell type BLANK");
								break;
							} // END SWITCH							
						} // END of IF CELL
						else value = "";
					} // END of IF ROW
					else value = "";
				} // END of IF SHEET
				else value = "";
			} // END of IF WORKBOOK
			else value = "";
		
		}
		catch (Exception e)
		{
			throw new ServiceException(e.getMessage());
		}
		
		// pipeline
		IDataCursor pipelineCursor_1 = pipeline.getCursor();
		IDataUtil.put( pipelineCursor_1, "cell_data", value );
		pipelineCursor_1.destroy();
		
		
		// --- <<IS-END>> ---

                
	}

	// --- <<IS-START-SHARED>> ---
	private static final long   DAY_MILLISECONDS  = 24 * 60 * 60 * 1000;
	
		////////////////////////////////////
		// Debug method	
		protected static Properties _props;
	
		private static boolean DEBUG = false;
		private static void printdbg(String msg)
		{
			String str = getProperty("enableDebug", "false");
			
			Boolean dbg = Boolean.valueOf(str); 
			if (dbg.booleanValue())
				System.out.println(msg);
		}
		////////////////////////////////////
	
	    protected static Properties getProps()
	    {
	        if(_props == null)
	            try
	            {
	                File cfgfn = new File(Server.getResources().getPackageConfigDir("SimpleExcel"), "excel.cnf");
	                if(cfgfn.exists())
	                {
	                    Properties tmp = new Properties();
	                    FileInputStream fin = new FileInputStream(cfgfn);
	                    tmp.load(fin);
	                    fin.close();
	                    _props = tmp;
	                }
	            }
	            catch(IOException io) { }
	        return _props;
	    }
	
	    public static String getProperty(String propertyName, String defValue)
	    {
	        Properties props = getProps();
	        String retval = null;
	        if(props != null)
	            retval = props.getProperty(propertyName, defValue);
	        return retval;
	    }
	
	public static String build_name(String str)
	{
	  String name = "";
	  StringTokenizer strtok = new StringTokenizer(str," ");
	  String temp = "";
	  int count = 0;
	  while (strtok.hasMoreElements())
	  {
	    temp = (String)strtok.nextElement();
	    if (count == 0)
	        name = temp;
	    else name += "_"+temp;
	
	    count++;
	  }
	  return name;
	}
	      /**
	       * Given a double, checks if it is a valid Excel date.
	       *
	       * @return true if valid
	       * @param  value the double value
	       */
	      public static boolean isValidExcelDate(double value)
	      {
	          return (value > -Double.MIN_VALUE);
	      }
	
	  ///////////////////////////////////////////////////////////////
	  // Method returns Java date pattern mapped from Excel date
	  public static String getCellDateFormat(HSSFCell cell)
	  {
		  String dt = "dd-MMM-yy";
	      HSSFCellStyle style = cell.getCellStyle();
	      int i = style.getDataFormat();
	      switch(i) 
	      {
	    // Internal Date Formats as described on page 427 in Microsoft Excel Dev's Kit...
	        case 0x0e: //m/d/yyyy
	        	dt = "MM/dd/yyyy";
	        	break;
	        case 0x0f: //d-mmm
	        	dt = "dd-MMM";
	        	break;
	        case 0x10: //d-mmm-yy
	        	dt = "dd-MMM-yy";
	        	break;
	        case 0x11: //mmm-yy
	        	dt = "MMMM-yy";
	        	break;
	        case 0x12: //h:mmAM/PM
	        	dt = "hh:mm aa";
	        	break;        	
	        case 0x13: //h:mm:ssAM/PM
	        	dt = "hh:mm:ss aa";
	        	break;
	        case 0x14: //h:mm
	        	dt = "hh:mm";
	        	break;
	        case 0x15: //h:mm:ss
	        	dt = "hh:mm:ss";
	        	break;
	        case 0x16: //m/d/yyyy h:mm
	        	dt = "MM/dd/yyyy hh:mm";
	        	break;
	        case 0x2d: //mm:ss
	        	dt = "mm:ss";
	        	break;
	        case 0x2e: //[h]:mm:ss
	        	dt = "hh:mm:ss";
	        	break;
	        case 0x2f: //mm:ss.0
	            dt = "mm:ss.SSSS";
	        break;
	        
	        default:
	        	dt = "dd-MMM-yy";
	        break;
	      }  
		  
		  return dt;
	  }
	
	//////////////////////////////////////////////////////////////////
	// method to determine if the cell is a date, versus a number...
	public static boolean isCellDateFormatted(HSSFCell cell) 
	{
	    boolean bDate = false;
	
	    double d = cell.getNumericCellValue();
	    if ( isValidExcelDate(d) ) {
	      HSSFCellStyle style = cell.getCellStyle();
	      int i = style.getDataFormat();
	      switch(i) {
	    // Internal Date Formats as described on page 427 in Microsoft Excel Dev's Kit...
	        case 0x0e:
	        case 0x0f:
	        case 0x10:
	        case 0x11:
	        case 0x12:
	        case 0x13:
	        case 0x14:
	        case 0x15:
	        case 0x16:
	        case 0x2d:
	        case 0x2e:
	        case 0x2f:
	         bDate = true;
	        break;
	
	        default:
	         bDate = false;
	        break;
	      }
	    }
	    return bDate;
	  }
	
	      /**
	       * Given a Calendar, return the number of days since 1600/12/31.
	       *
	       * @return days number of days since 1600/12/31
	       * @param  cal the Calendar
	       * @exception IllegalArgumentException if date is invalid
	       */
	
	      private static int absoluteDay(Calendar cal)
	      {
	          return cal.get(Calendar.DAY_OF_YEAR)
	                 + daysInPriorYears(cal.get(Calendar.YEAR));
	      }
	
	      /**
	       * Return the number of days in prior years since 1601
	       *
	       * @return    days  number of days in years prior to yr.
	       * @param     yr    a year (1600 < yr < 4000)
	       * @exception IllegalArgumentException if year is outside of range.
	       */
	
	      private static int daysInPriorYears(int yr)
	      {
	          if (yr < 1601)
	          {
	              throw new IllegalArgumentException(
	                  "'year' must be 1601 or greater");
	          }
	          int y    = yr - 1601;
	          int days = 365 * y      // days in prior years
	                     + y / 4      // plus julian leap days in prior years
	                     - y / 100    // minus prior century years
	                     + y / 400;   // plus years divisible by 400
	
	          return days;
	      }
	
	      /**
	       *  Given an Excel date with either 1900 or 1904 date windowing,
	       *  converts it to a java.util.Date.
	       *
	       *  @param date  The Excel date.
	       *  @param use1904windowing  true if date uses 1904 windowing,
	       *   or false if using 1900 date windowing.
	       *  @return Java representation of the date, or null if date is not a valid Excel date
	       */
	      public static Date getJavaDate(double date, boolean use1904windowing) {
	          if (isValidExcelDate(date)) {
	              int startYear = 1900;
	              int dayAdjust = -1; // Excel thinks 2/29/1900 is a valid date, which it isn't
	              int wholeDays = (int)Math.floor(date);
	              if (use1904windowing) {
	                  startYear = 1904;
	                  dayAdjust = 1; // 1904 date windowing uses 1/2/1904 as the first day
	              }
	              else if (wholeDays < 61) {
	                  // Date is prior to 3/1/1900, so adjust because Excel thinks 2/29/1900 exists
	                  // If Excel date == 2/29/1900, will become 3/1/1900 in Java representation
	                  dayAdjust = 0;
	              }
	              GregorianCalendar calendar = new GregorianCalendar(startYear,0, wholeDays + dayAdjust);
	              int millisecondsInDay = (int)((date - Math.floor(date)) * (double) DAY_MILLISECONDS + 0.5);
	              calendar.set(GregorianCalendar.MILLISECOND, millisecondsInDay);
	              return calendar.getTime();
	          }
	          else {
	              return null;
	          }
	      }
	// --- <<IS-END-SHARED>> ---
}

