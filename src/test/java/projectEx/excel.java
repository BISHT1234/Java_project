package projectEx;

import org.apache.poi.ss.usermodel.*;

import java.io.IOException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.Date;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.*;
public class excel {
   static	XSSFWorkbook workbook;
	static XSSFSheet sheet;
	excel(String epath,String shet){
	try {	 workbook = new XSSFWorkbook(epath);
	 sheet=workbook.getSheet(shet);}
	catch(Exception e) {
		System.out.print(e.getMessage());
	}
	}
	 
public static void who_Worked_For_7DAys() throws ParseException {
	System.out.print("employees  who has worked for 7 consecutive days.");
	try {
		//String excelPath="./data/Assignment_Timecard.xlsx";
		//XSSFWorkbook workbook = new XSSFWorkbook(excelPath);
		//XSSFSheet sheet=workbook.getSheet("sheet1");
		Iterator <Row>rowiterator =sheet.iterator();
	Row headerRow= rowiterator.next();
		DataFormatter formatter = new DataFormatter();
		String name=null;
		int count=0,previousdate=0;
		String pid=null;
while(rowiterator.hasNext()) {
	Row row=rowiterator.next();
	Cell cellname=row.getCell(7);
	String employeename=cellname.getStringCellValue();
    Cell celldate=row.getCell(2);
    Cell cellid=row.getCell(0);
  String pID=cellid.getStringCellValue();
	String value=formatter.formatCellValue(celldate);
	CellType type=celldate.getCellType();
	String data=value;
	if(data!="") {
	DateTimeFormatter inputFormat =	DateTimeFormatter.ofPattern("MM/dd/yyyy hh:mm a");
		LocalDateTime date = LocalDateTime.parse(value,inputFormat);
		DateTimeFormatter outputFormat = DateTimeFormatter.ofPattern("dd");
	
	String Date=date.format(outputFormat);
int newDate=Integer.parseInt(Date);
	       if(previousdate!=newDate) {
		        	if(name==employeename) {
		        	
		        		if(newDate==(previousdate+1)) {
		        			count++;
		        		
		        		}
		        		else {
		        			count=0;
		        		}
		        		previousdate=newDate;
		        	}
		        	else {
		        		previousdate=newDate;
		        		name=employeename;
		        		pid=pID;
		        	count=0;
		        	}
		        }
		 
	}
    if(count>=7) {
     	System.out.println("Name: "+employeename+"  position: "+pid);
     	  count=0;} 

	
}
	}catch(Exception exp) {
		System.out.print(exp);
	}
}   




public static void who_work_for_more_than_14hours() throws ParseException, IOException {
	System.out.println("Employees Who has worked for more than 14 hours in a single shift");
	String excelPath="./data/Assignment_Timecard.xlsx";
	XSSFWorkbook workbook = new XSSFWorkbook(excelPath);
	XSSFSheet sheet=workbook.getSheet("sheet1");
	Iterator <Row>rowiterator =sheet.iterator();
Row headerRow= rowiterator.next();
DataFormatter formatter = new DataFormatter();
String name=null;
int count=0,previousdate=0,maxtime=0;
String pid=null;
while(rowiterator.hasNext()) {
	Row row=rowiterator.next();
	Cell cellname=row.getCell(7);
	  Cell celldate=row.getCell(2);
	  Cell cellid=row.getCell(0);
	  String pID=cellid.getStringCellValue();
	  String value=formatter.formatCellValue(celldate);
	String employeename=cellname.getStringCellValue();
	Cell celltimecard=row.getCell(4);
String wtime=celltimecard.getStringCellValue();
String data=value;
if(data!="") {
DateTimeFormatter inputFormat =	DateTimeFormatter.ofPattern("MM/dd/yyyy hh:mm a");
	LocalDateTime date = LocalDateTime.parse(value,inputFormat);
	DateTimeFormatter outputFormat = DateTimeFormatter.ofPattern("dd");
  String Date=date.format(outputFormat);
int newDate=Integer.parseInt(Date);
String sub;
int i=0;
if(wtime!="") {
	sub=wtime.substring(0,1);
	i=Integer.parseInt(sub);}
	
	//System.out.println(i);
	  if(previousdate==newDate){
    	maxtime= count+i;  
    	name=employeename;
    	pid=pID;
	  }
       else {
    	   previousdate=newDate;
    	   count=i;
       }
}
if(maxtime>14) {
	System.out.println("Name: "+name+" position: "+pid);
}

}
}




public static void  who_have_less_than_10_hours() throws IOException {
	System.out.println("employees  who have less than 10 hours of time between shifts but greater than 1 hour");
	String excelPath="./data/Assignment_Timecard.xlsx";
	XSSFWorkbook workbook = new XSSFWorkbook(excelPath);
	XSSFSheet sheet=workbook.getSheet("sheet1");
	Iterator <Row>rowiterator =sheet.iterator();
Row headerRow= rowiterator.next();
DataFormatter formatter = new DataFormatter();
String name=null;
int count=0,previousdate=0,maxtime=0;
String pid=null,pname=null;
LocalDateTime pdate=null;
while(rowiterator.hasNext()) {
	Row row=rowiterator.next();
	Cell cellname=row.getCell(7);
	  Cell celldate=row.getCell(2);
	  Cell cellend=row.getCell(3);
	  Cell cellid=row.getCell(0);
	  String pID=cellid.getStringCellValue();
	  String value=formatter.formatCellValue(celldate);
	  String endtime=formatter.formatCellValue(cellend);
	String employeename=cellname.getStringCellValue();
String data=value;
if(data!=""&&endtime!="") {
DateTimeFormatter inputFormat =	DateTimeFormatter.ofPattern("MM/dd/yyyy hh:mm a");
	LocalDateTime date = LocalDateTime.parse(value,inputFormat);
	LocalDateTime date2 = LocalDateTime.parse(endtime,inputFormat);
	DateTimeFormatter outputFormat = DateTimeFormatter.ofPattern("hh:mm:ss a");
if(pdate!=null&&pname==employeename) {
	  Duration duration=Duration.between(pdate, date);
	//int newDate=Integer.parseInt(Date);
long d=duration.toHours();
if(d<10&&d>1) {
	System.out.println("Name: "+employeename+"  position: "+pID);
}
}
  pdate=date2;
pname=employeename;
}


}
	
}
}
