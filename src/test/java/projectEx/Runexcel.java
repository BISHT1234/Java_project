package projectEx;

import java.io.IOException;
import java.text.ParseException;

public class Runexcel {
	public static void main(String[] args) throws ParseException, IOException {
		String path="./data/Assignment_Timecard.xlsx";
		String sheet="sheet1";
		 excel ex=new excel(path,sheet);
		 ex.who_Worked_For_7DAys();
		 ex.who_work_for_more_than_14hours();
		 ex.who_have_less_than_10_hours();
	}
}
