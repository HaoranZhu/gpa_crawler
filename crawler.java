

package codes;

import java.io.*;
import java.sql.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;

import org.apache.poi.*;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class crawler {
	public static void main(String args[]) throws IOException{
		String classNameArr[]= new String[100];
		int creditArr[]=new int[100];
		String grade[]=new String[100];
		double result[]=new double[100];
		
		System.out.println("start gpa crawler!");
		String filePath = "/Users/Day1/Desktop/gpa.txt";
		File file=new File(filePath);
        if(file.isFile() && file.exists()){ //判断文件是否存在
            InputStreamReader read = new InputStreamReader(new FileInputStream(file));
            BufferedReader bufferedReader = new BufferedReader(read);
            String lineTxt = null;
            int cnt=0;
            int cnt2=1;
            while((lineTxt = bufferedReader.readLine()) != null){  	
            	
                int p0=lineTxt.indexOf("main");
                if(p0!=-1){
                lineTxt=lineTxt.substring(p0+8);
                int p1=lineTxt.indexOf("<");
                String className=lineTxt.substring(0, p1);
                //System.out.print(cnt+"	"+className);
                classNameArr[cnt]=className;
                cnt2++;
                }	    
                int p2=lineTxt.indexOf("</td><td>");
                if(p2!=-1){
                	lineTxt=lineTxt.substring(p2+9);
                	if(cnt2==2){
                		int p4=lineTxt.indexOf("<td>");
                		lineTxt=lineTxt.substring(p4+4,p4+5);
                		creditArr[cnt]=Integer.valueOf(lineTxt);
                		//System.out.print("\t"+lineTxt);
                	}
                	if(cnt2!=2&&cnt2!=3&&cnt2!=4){
                		if(grade[cnt].equals("P")){
                			result[cnt]=0;
                		}
                		else{
                			result[cnt]=Double.valueOf(lineTxt);
                		}
                		//System.out.print("\t"+lineTxt);
                	}
                	cnt2++;
                }
                
                int p3=lineTxt.indexOf("</td><td style=");
                if(p3!=-1){
                	lineTxt=lineTxt.substring(p3+19);
                	if(cnt2!=3&&cnt2!=4){
                		grade[cnt]=lineTxt;
                		//System.out.print("\t"+lineTxt);
                	}
                	cnt2++;
                }
                if(cnt2==7){
                	cnt++;
                	System.out.println();
                	cnt2=1;
                }
            }
            read.close();
		}else{
		    System.out.println("找不到指定的文件");
		}
        
        for(int i=0;i<50;i++){
        	System.out.println(i+"\t"+classNameArr[i]+"\t"+creditArr[i]+"\t"+grade[i]+"\t"+result[i]);
        }
        System.out.println("数据已经输入完毕");
        
        
        HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet("Sample sheet");	
		Map<String, Object[]> data = new HashMap<String, Object[]>();
		data.put("1", new Object[] {"ClassName","credit","grade","result"});
		for(int i=0;i<50;i++){
			int row=i+2;
			String rowString=String.valueOf(row);
			data.put(rowString, new Object[] {classNameArr[i],String.valueOf(creditArr[i]),grade[i],String.valueOf(result[i])});
			
			
		}
		/*
		data.put("1", new Object[] {"Emp No.", "Name", "Salary"});
		data.put("2", new Object[] {1d, "John", 1500000d});
		data.put("3", new Object[] {2d, "Sam", 800000d});
		data.put("4", new Object[] {3d, "Dean", 700000d});*/
		
		
		
		Set<String> keyset = data.keySet();
		int rownum = 0;
		for (String key : keyset) {
			Row row = sheet.createRow(rownum++);
			Object [] objArr = data.get(key);
			int cellnum = 0;
			for (Object obj : objArr) {
				Cell cell = row.createCell(cellnum++);
				if(obj instanceof Date) 
					cell.setCellValue((Date)obj);
				else if(obj instanceof Boolean)
					cell.setCellValue((Boolean)obj);
				else if(obj instanceof String)
					cell.setCellValue((String)obj);
				else if(obj instanceof Double)
					cell.setCellValue((Double)obj);
			}
		}
		
		try {
			FileOutputStream out = 
					new FileOutputStream(new File("/Users/Day1/Desktop/exectTest.xls"));
			workbook.write(out);
			out.close();
			System.out.println("Excel written successfully..");
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

     
	}
	
}
