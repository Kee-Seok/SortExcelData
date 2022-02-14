package t;

import java.awt.Desktop;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;

import jxl.LabelCell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class Excel {

	static File ansysFile = new File("./temp/ansys.xls");
	static Workbook ansysWb;
	static Sheet ansysSheet;
	static ArrayList<String> ansysNameArray = new ArrayList<>();
	static ArrayList<String> ansysBirthArray = new ArrayList<>();
	static ArrayList<String> ansysNumArray = new ArrayList<>();
	static ArrayList<String> ansysPhoneNumbArray = new ArrayList<>();
	
	static File juminFile = new File("./temp/jumin.xls");
	static Workbook juminWb;
	static Sheet juminSheet;
	static ArrayList<String> juminNameArray = new ArrayList<>();
	static ArrayList<String> juminBirthArray = new ArrayList<>();
	static int row;
	
	public static void getDataFromAnsys() {
	try {
		ansysWb = Workbook.getWorkbook(ansysFile);
		ansysSheet = ansysWb.getSheet(0);
		for(int i = 1; i < ansysSheet.getRows(); i++) {
		ansysNameArray.add(ansysSheet.getCell(1,i).getContents().toString());
		ansysBirthArray.add(ansysSheet.getCell(2,i).getContents().toString());
		ansysNumArray.add(ansysSheet.getCell(0,i).getContents().toString());
		ansysPhoneNumbArray.add(ansysSheet.getCell(3,i).getContents().toString());
		}
	} catch (BiffException e) {
		e.printStackTrace();
	} catch (IOException e) {
		e.printStackTrace();
	}
	}
	
	public static void getDataFromJumin() {
	try {
		juminWb = Workbook.getWorkbook(juminFile);
		juminSheet = juminWb.getSheet(0);
		for(int i = 1; i < juminSheet.getRows(); i++) {
		juminNameArray.add(juminSheet.getCell(1,i).getContents().toString());
		juminBirthArray.add(juminSheet.getCell(2,i).getContents().toString());
		}
	} catch (BiffException e) {
		e.printStackTrace();
	} catch (IOException e) {
		e.printStackTrace();
	}
	}
	
	public static void writeExcel() throws IOException, RowsExceededException, WriteException {
		getDataFromAnsys();
		getDataFromJumin();
			File file = new File("./temp/SortedExcel.xls");
			file.createNewFile();
			WritableWorkbook wb = Workbook.createWorkbook(file);
			WritableSheet ws = wb.createSheet("실제명단",0);
			System.out.println(ansysSheet.getRows());
			for(int i = 0; i < ansysSheet.getRows()-1; i++) {
				ansysNameArray.get(i);
				ansysBirthArray.get(i);
				for(int j = 0; j < juminSheet.getRows()-1; j++) {
					juminNameArray.get(j);
					juminBirthArray.get(j);
					if(ansysNameArray.get(i).equals(juminNameArray.get(j))
					&&ansysBirthArray.get(i).equals(juminBirthArray.get(j))) {
						String[] str = {ansysNumArray.get(i), ansysNameArray.get(i),ansysBirthArray.get(i), ansysPhoneNumbArray.get(i)};
						for(int s = 0; s < str.length; s++) {
							ws.addCell(new Label(s,row,str[s]));
						}
						row++;
					}
				}
			}
			wb.write();
			wb.close();
	
	}
	
	public static void main(String[] args) {
		try {
			writeExcel();
			Desktop.getDesktop().open(new File("./temp/SortedExcel.xls"));
		} catch (RowsExceededException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
