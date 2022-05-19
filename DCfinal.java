
//FINAL CODE

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;
import java.io.*;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
class Student{
    String lastName;//Stores student names
    String firstName;
    String preference1;//Stores preferences from the preference sheet
    String preference2;
    String preference3;
    String preference4;
    String preference5;
    
    
    String period1;//Stores classes from the sample schedule sheet
    String period2;
    String period3;
    String period4;
    String period5;
    String period6;
    
    int unhappiness=0;//goes up when a student doesn't get a class they want
    
    
    
}
//This is a key for what the output for what the preference sorter looks like.
//A= rhetoric
//C= calc
//La= language
//Lc= coding
//P= physics

public class DaCoda_ScheduleRater {
    public static void scheduleRater() {
        try {
            String filename="Schedule.xlsx";//sets up excel reader
            FileInputStream fileInputStream = new FileInputStream(filename);
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
            XSSFSheet worksheet = workbook.getSheet("ExcelReader");//looks at preference sheet
            int rowTotal = worksheet.getPhysicalNumberOfRows();
            int rows = rowTotal;
            Student grade[] = new Student[rowTotal];
        for (int i = 0; i < rowTotal; i++){//reads excel  preference sheet
            XSSFRow row1 = worksheet.getRow(i);
            XSSFCell cellA1 = row1.getCell((short) 0);
            String a1Val = cellA1.getStringCellValue();
            XSSFCell cellB1 = row1.getCell((short) 1);
            String b1Val = cellB1.getStringCellValue();
            XSSFCell cellC1 = row1.getCell((short) 2);
            String c1Val = cellC1.getStringCellValue();
            XSSFCell cellD1 = row1.getCell((short) 3);
            String d1Val = cellD1.getStringCellValue();
            XSSFCell cellE1 = row1.getCell((short) 4);
            String e1Val = cellE1.getStringCellValue();
            XSSFCell cellF1 = row1.getCell((short) 5);
            String f1Val = cellF1.getStringCellValue();
            XSSFCell cellG1 = row1.getCell((short) 6);
            String g1Val = cellG1.getStringCellValue();


            grade[i]=new Student();//assigns data to student object
            grade[i].lastName= a1Val;
            grade[i].firstName=b1Val;
            grade[i].preference1=c1Val;
            grade[i].preference2=d1Val;
            grade[i].preference3=e1Val;
            grade[i].preference4=f1Val;
            grade[i].preference5=g1Val;
            
            
            }
        XSSFSheet worksheet2 = workbook.getSheet("Scheduling One");//Reads sample schedule sheet
        int rowTotal2 = worksheet2.getPhysicalNumberOfRows();
        
        for(int i=0;i<rowTotal;i++){//reads students assigned classes
            XSSFRow row1 = worksheet2.getRow(i);
            
            XSSFCell cellC1 = row1.getCell((short) 2);
            String c1Val = cellC1.getStringCellValue();
            XSSFCell cellD1 = row1.getCell((short) 3);
            String d1Val = cellD1.getStringCellValue();
            XSSFCell cellE1 = row1.getCell((short) 4);
            String e1Val = cellE1.getStringCellValue();
            XSSFCell cellF1 = row1.getCell((short) 5);
            String f1Val = cellF1.getStringCellValue();
            XSSFCell cellG1 = row1.getCell((short) 6);
            String g1Val = cellG1.getStringCellValue();
            XSSFCell cellH1 = row1.getCell((short) 7);
            String h1Val = cellH1.getStringCellValue();
            
            grade[i].period1= c1Val;//assigns saves read data
            grade[i].period2=d1Val;
            grade[i].period3=e1Val;
            grade[i].period4=f1Val;
            grade[i].period5=g1Val;
            grade[i].period6=h1Val;
            
        }
        /*
         * Each student is given a different unhappiness value, the more preferred
         * classes a student does not receiv=-9821q, the higher the unhappiness value is.
         */
        for(int i=0;i<rowTotal;i++){ //compares the students' preferences to their assigned classes 
            if(grade[i].preference1.equals(grade[i].period1)||grade[i].preference1.equals(grade[i].period2)||grade[i].preference1.equals(grade[i].period3)||grade[i].preference1.equals(grade[i].period4)||grade[i].preference1.equals(grade[i].period5)||grade[i].preference1.equals(grade[i].period6)){
                grade[i].unhappiness+=0;
            }
            else{
                grade[i].unhappiness+=1;
            }
            if(grade[i].preference2.equals(grade[i].period1)||grade[i].preference2.equals(grade[i].period2)||grade[i].preference2.equals(grade[i].period3)||grade[i].preference2.equals(grade[i].period4)||grade[i].preference2.equals(grade[i].period5)||grade[i].preference2.equals(grade[i].period6)){
                grade[i].unhappiness+=0;
            }
            else{
                grade[i].unhappiness+=1;
            }
            if(grade[i].preference3.equals(grade[i].period1)||grade[i].preference3.equals(grade[i].period2)||grade[i].preference3.equals(grade[i].period3)||grade[i].preference3.equals(grade[i].period4)||grade[i].preference3.equals(grade[i].period5)||grade[i].preference3.equals(grade[i].period6)){
                grade[i].unhappiness+=0;
            }
            else{
                grade[i].unhappiness+=1;
            }
        }
        int allWantCount=0;//number of people got all preferences
        int twoClassCount=0;//number of peopel that got two preferences
        int oneClassCount=0;//number of people that got one preference
        int noClassCount=0;//number of people that got no preferences
        
        for(int i=0;i<rowTotal;i++){ //counts stats based on level of unhappiness
            if(grade[i].unhappiness==0){
                allWantCount++;
            }
            else if(grade[i].unhappiness==1){
                twoClassCount++;
            }
            else if(grade[i].unhappiness==2){
                oneClassCount++;
            }
            else if(grade[i].unhappiness==3){
                noClassCount++;
            }
        }
        //prints data to the terminal as well as the excel sheet
        System.out.println("These many people got all there classes: "+allWantCount);
        System.out.println("These many people got two of there classes: "+twoClassCount);
        System.out.println("These many people got one of there classes: "+oneClassCount);
        System.out.println("These many people got none of there classes: "+noClassCount);
        
        System.out.println("People who only got two preference:");
        String[] twoClassPeople = new String[rowTotal];//puts names into its own array
        for(int i=0;i<rowTotal;i++){
            if(grade[i].unhappiness==1){
                System.out.println(grade[i].firstName+" "+grade[i].lastName);
                twoClassPeople[i]=grade[i].firstName+" "+grade[i].lastName;
            }
        }
        System.out.println("People who only got one preference:");
        String[] oneClassPeople = new String[rowTotal];//array for one preference people
        for(int i=0;i<rowTotal;i++){
            if(grade[i].unhappiness==2){
                System.out.println(grade[i].firstName+" "+grade[i].lastName);
                oneClassPeople[i]=grade[i].firstName+" "+grade[i].lastName;
            }
        }
        System.out.println("People who got no preference:");
        String[] noClassPeople = new String[rowTotal];//array for people with no preferences
        for(int i=0;i<rowTotal;i++){
            if(grade[i].unhappiness==3){
                System.out.println(grade[i].firstName+" "+grade[i].lastName);
                noClassPeople[i]=grade[i].firstName+" "+grade[i].lastName;
            }
        }
        
        XSSFSheet worksheet3 = workbook.getSheet("Data");//prints stats to excel
        XSSFRow row = worksheet3.createRow((short)0);
        row.createCell(0).setCellValue("Students with all preferences");
        row.createCell(1).setCellValue("Students with two preferences");
        row.createCell(2).setCellValue("Students with one preference");
        row.createCell(3).setCellValue("Students with no preferences");
        

        XSSFRow rowhead = worksheet3.createRow((short)1);
                
    
        rowhead.createCell(0).setCellValue(allWantCount);
        rowhead.createCell(1).setCellValue(twoClassCount);
        rowhead.createCell(2).setCellValue(oneClassCount);
        rowhead.createCell(3).setCellValue(noClassCount);
        int twoRowCount=3;
        int oneRowCount=3;
        int noRowCount=3;
        
        for(int i=3;i<rowTotal;i++){ //creates all rows needed to fill data
            XSSFRow roworiginal = worksheet3.createRow((short)i);
        }
        for(int i=0;i<rowTotal;i++){ //fills data into created rows
            
            if(twoClassPeople[i]!=null||oneClassPeople[i]!=null||noClassPeople[i]!=null){ //checks to make sure there is data to display
                /*
                 * The names of the students are added if their is a name at the given
                 * index in each array (if the array index is empty it does not print anything)
                 */
            
                if(twoClassPeople[i]!=null){
                XSSFRow row1 = worksheet3.getRow((short)twoRowCount);
                    Cell cell1 = row1.createCell(1);
                    cell1.setCellValue((twoClassPeople[i]));
                    twoRowCount++;
                }
                if(oneClassPeople[i]!=null){
                XSSFRow row1 = worksheet3.getRow((short)oneRowCount);
                    Cell cell2 = row1.createCell(2);
                    cell2.setCellValue((oneClassPeople[i]));
                    oneRowCount++;
                }
                if(noClassPeople[i]!=null){
                XSSFRow row1 = worksheet3.getRow((short)noRowCount);
                    Cell cell3= row1.createCell(3);
                    cell3.setCellValue((noClassPeople[i]));
                    noRowCount++;
                }
                
            
            }
        }
        
        FileOutputStream fileOut = new FileOutputStream(filename);
        workbook.write(fileOut);//writes the data to excel
        //closing the Stream
        fileOut.close();
        //closing the workbook
        workbook.close();
        } catch (FileNotFoundException e) {
        e.printStackTrace();
        } catch (IOException e) {
        e.printStackTrace();
        }
    
    
    }
    public static void DaCoda(){
        Scanner keys = new Scanner(System.in);
        String a1Val = "x";
        String b1Val = "x";
        String c1Val = "x";
        String d1Val = "x";
      /* Key for sorter
         * A-Rhetoric
         * C- Calculus
         * Lc- Logic and Coding
         * La- Language
         * P-physics
         */

        int TraditionalNumber = 0;
        int AllMathScience = 0;
        int NoMath = 0;
        int ACLa = 0;
        int ACLc = 0;
        int ACP = 0;
        int ALcP = 0;
        int CLaLc = 0;
        int LaLcP = 0;
        try{    
            //declare file name to be created.
            //"C://Users//bdnel//Downloads//Schedule7.xlsx"
            String filename = "Schedule.xlsx";
            FileInputStream fileInputStream = new FileInputStream(filename);
            //creating an instance of XSSFWorkbook class
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
            //invoking creatSheet() method and passing the name of the sheet to be created
            XSSFSheet sheet = workbook.getSheet("ExcelReader");
            //creating the 0th row using the createRow() method
            int rowTotal = sheet.getPhysicalNumberOfRows();
            //creating cell by using the createCell() method and setting the values to the cell by using the setCellValue() method
            //creating the 1st row
            Student[] grade = new Student[rowTotal];
            for(int i = 0; i<grade.length; i++){
                //System.out.println(i);
                XSSFRow row1 = sheet.getRow(i);//reads data from spreadsheet
                XSSFCell cellA1 = row1.getCell((short) 2);
                a1Val = cellA1.getStringCellValue();
                XSSFCell cellB1 = row1.getCell((short) 3);
                b1Val = cellB1.getStringCellValue();
                XSSFCell cellC1 = row1.getCell((short) 4);
                c1Val = cellC1.getStringCellValue();
                XSSFCell cellD1 = row1.getCell((short) 0);
                d1Val = cellD1.getStringCellValue();
                //System.out.println(a1Val);
                //System.out.println(b1Val);
                //System.out.println(c1Val);
                //System.out.println(s1.length);
                //System.out.println(rowTotal);
                grade[i] = new Student();//stores data into objects
                grade[i].preference1 = a1Val;
                grade[i].preference2 = b1Val;
                grade[i].preference3 = c1Val;
                grade[i].firstName = d1Val;
                String[] arr = {grade[i].preference1, grade[i].preference2, grade[i].preference3};
                int n = arr.length;
                for(int m = 0; m<n; m++){  //takes top 3 preferences and arranges them
                     for (int j = m+1; j<n; j++){  
                        //compares each elements of the array to all the remaining elements
                        //fruits is the name of the array
                        if(arr[m].compareTo(arr[j])>0){  
                            //swapping array elements  
                            String temp = arr[m];  
                            arr[m] = arr[j];  
                            arr[j] = temp;  
                        }  
                     }  
               }  
               grade[i].preference1 = arr[0];
               grade[i].preference2 = arr[1];
               grade[i].preference3 = arr[2];
//Checks combos from data and figures out how many of each combo there is.
               if((grade[i].preference1.matches("Calculus II"))&&(grade[i].preference2.matches("Language - Greek II"))&&(grade[i].preference3.matches("Physics II"))){
                   TraditionalNumber++;
                   //System.out.println("Yes");
               }
               if((grade[i].preference1.matches("Calculus II"))&&(grade[i].preference2.matches("Language - Spanish IV"))&&(grade[i].preference3.matches("Physics II"))){
                   TraditionalNumber++;
                   //System.out.println("Yes");
               }
               if((grade[i].preference1.matches("Calculus II"))&&(grade[i].preference2.matches("Language - French IV"))&&(grade[i].preference3.matches("Physics II"))){
                   TraditionalNumber++;
                   //System.out.println("Yes");
               }
               if((grade[i].preference1.matches("Calculus II"))&&(grade[i].preference2.matches("Language - German IV"))&&(grade[i].preference3.matches("Physics II"))){
                   TraditionalNumber++;
                   //System.out.println("Yes");
               }
               if((grade[i].preference1.matches("Calculus II"))&&(grade[i].preference2.matches("Logic and Coding"))&&(grade[i].preference3.matches("Physics II"))){
                  AllMathScience++;
                  //System.out.println("Yes Yes");
               }
               if((grade[i].preference1.matches("American Rhetorical Tradition"))&&(grade[i].preference2.matches("Language - Greek II"))&&(grade[i].preference3.matches("Logic and Coding"))){
                   NoMath++;
                  // System.out.println("Yes No");
               }
               if((grade[i].preference1.matches("American Rhetorical Tradition"))&&(grade[i].preference2.matches("Language - Spanish IV"))&&(grade[i].preference3.matches("Logic and Coding"))){
                   NoMath++;
                   //System.out.println("Yes No");
               }
               if((grade[i].preference1.matches("American Rhetorical Tradition"))&&(grade[i].preference2.matches("Language - French IV"))&&(grade[i].preference3.matches("Logic and Coding"))){
                   NoMath++;
                   //System.out.println("Yes No");
               }
               if((grade[i].preference1.matches("American Rhetorical Tradition"))&&(grade[i].preference2.matches("Language - German IV"))&&(grade[i].preference3.matches("Logic and Coding"))){
                   NoMath++;
                  // System.out.println("Yes No");
               }
               if((grade[i].preference1.matches("American Rhetorical Tradition"))&&(grade[i].preference2.matches("Calculus II"))&&(grade[i].preference3.matches("Language - Greek II"))){
                   ACLa++;
                   //System.out.println("Yes No 1");
               }
               if((grade[i].preference1.matches("American Rhetorical Tradition"))&&(grade[i].preference2.matches("Calculus II"))&&(grade[i].preference3.matches("Language - Spanish IV"))){
                   ACLa++;
                   //System.out.println("Yes No 1");
               }
               if((grade[i].preference1.matches("American Rhetorical Tradition"))&&(grade[i].preference2.matches("Calculus II"))&&(grade[i].preference3.matches("Language - French IV"))){
                   ACLa++;
                   //System.out.println("Yes No 1");
               }
               if((grade[i].preference1.matches("American Rhetorical Tradition"))&&(grade[i].preference2.matches("Calculus II"))&&(grade[i].preference3.matches("Language - German IV"))){
                   ACLa++;
                   //System.out.println("Yes No 1");
               }
               if((grade[i].preference1.matches("American Rhetorical Tradition"))&&(grade[i].preference2.matches("Calculus II"))&&(grade[i].preference3.matches("Logic and Coding"))){
                   ACLc++;
                   //System.out.println("Yes No 2");
               }
               if((grade[i].preference1.matches("American Rhetorical Tradition"))&&(grade[i].preference2.matches("Calculus II"))&&(grade[i].preference3.matches("Physics II"))){
                   ACP++;
                  // System.out.println("Yes No 3");
               }
               if((grade[i].preference1.matches("American Rhetorical Tradition"))&&(grade[i].preference2.matches("Logic and Coding"))&&(grade[i].preference3.matches("Physics II"))){
                   ALcP++;
                   //System.out.println("Yes No 4");
               }
               if((grade[i].preference1.matches("Calculus II"))&&(grade[i].preference2.matches("Language - Greek II"))&&(grade[i].preference3.matches("Logic and Coding"))){
                   NoMath++;
                  // System.out.println("Yes No 5");
               }
               if((grade[i].preference1.matches("Calculus II"))&&(grade[i].preference2.matches("Language - Spanish IV"))&&(grade[i].preference3.matches("Logic and Coding"))){
                   NoMath++;
                  // System.out.println("Yes No 5");
               }
               if((grade[i].preference1.matches("Calculus II"))&&(grade[i].preference2.matches("Language - French IV"))&&(grade[i].preference3.matches("Logic and Coding"))){
                   NoMath++;
                  // System.out.println("Yes No 5");
               }
               if((grade[i].preference1.matches("Calculus II"))&&(grade[i].preference2.matches("Language - German IV"))&&(grade[i].preference3.matches("Logic and Coding"))){
                   CLaLc++;
                  // System.out.println("Yes No 5");
               }
               if((grade[i].preference1.matches("Language - Greek II"))&&(grade[i].preference2.matches("Logic and Coding"))&&(grade[i].preference3.matches("Physics II"))){
                   LaLcP++;
                   //System.out.println("Yes No 7");
               }
               if((grade[i].preference1.matches("Language - Spanish IV"))&&(grade[i].preference2.matches("Logic and Coding"))&&(grade[i].preference3.matches("Physics II"))){
                   LaLcP++;
                   //System.out.println("Yes No 7");
               }
               if((grade[i].preference1.matches("Language - French IV"))&&(grade[i].preference2.matches("Logic and Coding"))&&(grade[i].preference3.matches("Physics II"))){
                   LaLcP++;
                  // System.out.println("Yes No 7");
               }
               if((grade[i].preference1.matches("Language - German IV"))&&(grade[i].preference2.matches("Logic and Coding"))&&(grade[i].preference3.matches("Physics II"))){
                   LaLcP++;
                  // System.out.println("Yes No 7");
               }
             //  System.out.println(grade[i].preference1);
              // System.out.println(grade[i].preference2);
               //System.out.println(grade[i].preference3);
               //System.out.println(grade[i].firstName);
            }
	// Prints results to terminal
           // System.out.println("Excel Car file has been generated successfully.");
            System.out.println("There are "+TraditionalNumber+" students who want the traditional path.");
            System.out.println("There are "+AllMathScience+" students who want the all math and science path.");
            System.out.println("There are "+NoMath+" students who want the no math or science path");
            System.out.println("Rhetoric, Calculus, Language - "+ ACLa);
            System.out.println("Rhetoric, Calculus, Logic and Coding - "+ ACLc);
            System.out.println("Rhetoric, Calculus, Physics - "+ ACP);
            System.out.println("Rhetoric, Logic and Coding, Physics - "+ ALcP);
            System.out.println("Calculus, Language, Logic and Coding - "+ CLaLc);
            System.out.println("Language, Logic and Coding, Physics - "+ LaLcP);
            sheet = workbook.getSheet("BreakDown");
            XSSFRow rowhead = sheet.createRow((short)0);// prints results to excel
            rowhead.createCell(0).setCellValue("Rhetoric, Calculus, Language");
            rowhead.createCell(1).setCellValue("Rhetoric, Calculus, Logic and Coding");
            rowhead.createCell(2).setCellValue("Rhetoric, Calculus, Physics");
            rowhead.createCell(3).setCellValue("NoMath(ALaLC)");
            rowhead.createCell(4).setCellValue("Rhetoric, Logic and Coding, Physics");
            rowhead.createCell(5).setCellValue("Calculus, Language, Logic and Coding");
            rowhead.createCell(6).setCellValue("AllMath(CLcP)");
            rowhead.createCell(7).setCellValue("TraditionalPath(CLaP)");
            rowhead.createCell(8).setCellValue("Language, Logic and Coding, Physics");
            XSSFRow row = sheet.createRow((short)1);
            row.createCell(0).setCellValue(ACLa);
            row.createCell(1).setCellValue(ACLc);
            row.createCell(2).setCellValue(ACP);
            row.createCell(3).setCellValue(NoMath);
            row.createCell(4).setCellValue(ALcP);
            row.createCell(5).setCellValue(CLaLc);
            row.createCell(6).setCellValue(AllMathScience);
            row.createCell(7).setCellValue(TraditionalNumber);
            row.createCell(8).setCellValue(LaLcP);
            FileOutputStream fileOut = new FileOutputStream(filename);
            workbook.write(fileOut);
            //closing the Stream
            fileOut.close();
            //closing the workbook
            workbook.close();
            //prints the message on the console
            System.out.println("Excel file has been generated successfully.");
        }
        catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        catch (IOException e) {
            e.printStackTrace();
        }
    }
    public static void main(String[]args){
        Scanner keys = new Scanner(System.in);
        System.out.println("Please input the file location of your inputs");
         //String filename = keys.next();
        System.out.println("Press 1 to sort student preferences");
        System.out.println("Press 2 to rate a schedule");
        int input=keys.nextInt();
        if(input==1){
            DaCoda();
        }
        else if(input==2){
            scheduleRater();
        }
        else{
            System.out.println("Invalid Input");
        }
    }
}
/*The following portion of code is not actually functional. 
*It was our best attempt at a code that created schedules.
*Due to various difficulties, as well as time constraints, we were unable to finish it.
*We left it here as a comment to be helpful for anyone in the future who wishes
*to successfully make a schedule creator program
*/
/* 
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;
import java.io.*;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
//This is the actual scheduling part of the code that is unfinished and needs to be completed by the next year Seniors
public class Scheduling extends Subject_Operation{
    String a2Val;
    String b2Val;
    String c2Val;
    String d2Val;
    String e2Val;
    String f2Val;
    String g2Val;
void readSorting(){
    try {
    FileInputStream fileInputStream = new FileInputStream("C:\\Users\\ishik\\Downloads\\SeniorSchedulingCopy.xlsx");
    XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
    XSSFSheet worksheet = workbook.getSheet("Input Template (Class Availability)");
    int rowTotal = worksheet.getPhysicalNumberOfRows();
    Student grade[] = new Student[rowTotal];
    for (int i = 0; i < rowTotal; i++){
            grade[i] = new Student();
            XSSFRow row1 = worksheet.getRow(i);
            XSSFCell cellA1 = row1.getCell((short) 0);
            String a1Val = cellA1.getStringCellValue();
            XSSFCell cellB1 = row1.getCell((short) 1);
            String b1Val = cellB1.getStringCellValue();
            XSSFCell cellC1 = row1.getCell((short) 2);
            String c1Val = cellC1.getStringCellValue();
            XSSFCell cellD1 = row1.getCell((short) 3);
            String d1Val = cellD1.getStringCellValue();
            XSSFCell cellE1 = row1.getCell((short) 4);
            String e1Val = cellE1.getStringCellValue();
            XSSFCell cellF1 = row1.getCell((short) 5);
            String f1Val = cellF1.getStringCellValue();
            XSSFCell cellG1 = row1.getCell((short) 6);
            String g1Val = cellG1.getStringCellValue();
    }
    
    XSSFSheet sheet = workbook.getSheet("Input Template (Class Availability)");
    int rowTotal2 =  sheet.getPhysicalNumberOfRows();
    Class_Availability period = new Class_Availability();
    for (int i = 1; i < rowTotal2; i++){
    XSSFRow row2 = sheet.getRow(i);
    XSSFCell cellA2 = row2.getCell((short) 0);
    String a2Val = cellA2.getStringCellValue(); 
    period.firstPeriod[i] = a2Val;
    XSSFCell cellB2 = row2.getCell((short) 1);
    String b2Val = cellB2.getStringCellValue();
    period.secondPeriod[i] = b2Val;
    XSSFCell cellC2 = row2.getCell((short) 2);
    String c2Val = cellC2.getStringCellValue(); 
    period.thirdPeriod[i] = c2Val;
    XSSFCell cellD2 = row2.getCell((short) 3);
    String d2Val = cellD2.getStringCellValue();
    period.fourthPeriod[i] = d2Val;
    XSSFCell cellE2 = row2.getCell((short) 4);
    String e2Val = cellE2.getStringCellValue();
    period.fifthPeriod[i] = e2Val;
    XSSFCell cellF2 = row2.getCell((short) 5);
    String f2Val = cellF2.getStringCellValue();
    period.sixthPeriod[i] = f2Val;
    }
    
    System.out.println("Excel file has been read successfully");
    Class_Availability firstPeriod[] = new Class_Availability[rowTotal2];
    Class_Availability secondPeriod[] = new Class_Availability[rowTotal2];
    Class_Availability thirdPeriod[] = new Class_Availability[rowTotal2];
    Class_Availability fourthPeriod[] = new Class_Availability[rowTotal2];
    Class_Availability fifthPeriod[] = new Class_Availability[rowTotal2];
    Class_Availability sixthPeriod[] = new Class_Availability[rowTotal2];
    String art = "Art";
    String hl = "Humane Letters IV";
    String french = "Language - French IV";
    String greek = "Language - Greek II";
    String german = "Language - German IV";
    String spanish = "Language - Spanish IV";
    String coding = "Logic and Coding";
    String calc = "Calculus";
    String physics = "Physics IV";
    String rhetoric = "American Rhetorical Tradition";
    Subject availability = new Subject();
    for(int i = 0; i < grade.length; i++){
        for(int j = 0; j < firstPeriod.length; j++){
        if(firstPeriod[j].equals(grade[i].preference1) && (grade[i].preference1.equals(french) || grade[i].preference2.equals(french) || grade[i].preference3.equals(french))){
            grade[i].period1 = french;
            grade[i].period2 = art;
            for(int h = 0; h < thirdPeriod.length; h++){
                if(thirdPeriod[h].equals(grade[i].preference1) && (grade[i].preference1.equals(coding) || grade[i].preference2.equals(coding) || grade[i].preference3.equals(coding))){
                    grade[i].period3 = coding;
                    for(int k = 0; k < fourthPeriod.length; k++){
                        if(fourthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(rhetoric) || grade[i].preference2.equals(rhetoric) || grade[i].preference3.equals(rhetoric))){
                            grade[i].period1 = french;
                            grade[i].period2 = art;
                            grade[i].period3 = coding;
                            grade[i].period4 = rhetoric;
                            grade[i].period5 = hl;
                            grade[i].period6 = hl;
                        }
                        if(fourthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(physics) || grade[i].preference2.equals(physics) || grade[i].preference3.equals(physics))){
                            grade[i].period1 = french;
                            grade[i].period2 = art;
                            grade[i].period3 = coding;
                            grade[i].period4 = physics;
                            grade[i].period5 = hl;
                            grade[i].period6 = hl;
                        }
                        if(fourthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(calc) || grade[i].preference2.equals(calc) || grade[i].preference3.equals(calc))){
                            grade[i].period1 = french;
                            grade[i].period2 = art;
                            grade[i].period3 = coding;
                            grade[i].period4 = calc;
                            grade[i].period5 = hl;
                            grade[i].period6 = hl;
                        }
                    }
                }
                
                if(thirdPeriod[h].equals(grade[i].preference1) && (grade[i].preference1.equals(calc) || grade[i].preference2.equals(calc) || grade[i].preference3.equals(calc))){
                    grade[i].period3 = calc;
                    for(int k = 0; k < fourthPeriod.length; k++){
                        if(fourthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(rhetoric) || grade[i].preference2.equals(rhetoric) || grade[i].preference3.equals(rhetoric))){
                            grade[i].period1 = french;
                            grade[i].period2 = art;
                            grade[i].period3 = calc;
                            grade[i].period4 = rhetoric;
                            grade[i].period5 = hl;
                            grade[i].period6 = hl;
                        }
                        if(fourthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(physics) || grade[i].preference2.equals(physics) || grade[i].preference3.equals(physics))){
                            grade[i].period1 = french;
                            grade[i].period2 = art;
                            grade[i].period3 = calc;
                            grade[i].period4 = physics;
                            grade[i].period5 = hl;
                            grade[i].period6 = hl;
                        }
                        if(fourthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(coding) || grade[i].preference2.equals(coding) || grade[i].preference3.equals(coding))){
                            grade[i].period1 = french;
                            grade[i].period2 = art;
                            grade[i].period3 = calc;
                            grade[i].period4 = coding;
                            grade[i].period5 = hl;
                            grade[i].period6 = hl;
                        }
                    }
                }
                
                if(thirdPeriod[h].equals(grade[i].preference1) && (grade[i].preference1.equals(physics) || grade[i].preference2.equals(physics) || grade[i].preference3.equals(physics))){
                    grade[i].period3 = physics;
                    for(int k = 0; k < fourthPeriod.length; k++){
                        if(fourthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(rhetoric) || grade[i].preference2.equals(rhetoric) || grade[i].preference3.equals(rhetoric))){
                            grade[i].period1 = french;
                            grade[i].period2 = art;
                            grade[i].period3 = physics;
                            grade[i].period4 = rhetoric;
                            grade[i].period5 = hl;
                            grade[i].period6 = hl;
                        }
                        if(fourthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(calc) || grade[i].preference2.equals(calc) || grade[i].preference3.equals(calc))){
                            grade[i].period1 = french;
                            grade[i].period2 = art;
                            grade[i].period3 = physics;
                            grade[i].period4 = calc;
                            grade[i].period5 = hl;
                            grade[i].period6 = hl;
                        }
                        if(fourthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(coding) || grade[i].preference2.equals(coding) || grade[i].preference3.equals(coding))){
                            grade[i].period1 = french;
                            grade[i].period2 = art;
                            grade[i].period3 = physics;
                            grade[i].period4 = coding;
                            grade[i].period5 = hl;
                            grade[i].period6 = hl;
                        }
                    }
                }
                
                if(thirdPeriod[h].equals(grade[i].preference1) && (grade[i].preference1.equals(hl) || grade[i].preference2.equals(hl) || grade[i].preference3.equals(hl))){
                    grade[i].period3 = hl;
                    grade[i].period4 = hl;
                    for(int k = 0; k < fifthPeriod.length; k++){
                        for(int l = 0; l < sixthPeriod.length; l++){
                        if(fifthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(rhetoric) || grade[i].preference2.equals(rhetoric) || grade[i].preference3.equals(rhetoric))){
                            if(sixthPeriod[l].equals(grade[i].preference1) && (grade[i].preference1.equals(calc) || grade[i].preference2.equals(calc) || grade[i].preference3.equals(calc))){
                            grade[i].period1 = french;
                            grade[i].period2 = art;
                            grade[i].period3 = hl;
                            grade[i].period4 = hl;
                            grade[i].period5 = rhetoric;
                            grade[i].period6 = calc;
                            }
                        }
                        if(fifthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(calc) || grade[i].preference2.equals(calc) || grade[i].preference3.equals(calc))){
                            if(sixthPeriod[l].equals(grade[i].preference1) && (grade[i].preference1.equals(coding) || grade[i].preference2.equals(coding) || grade[i].preference3.equals(coding))){
                            grade[i].period1 = french;
                            grade[i].period2 = art;
                            grade[i].period3 = hl;
                            grade[i].period4 = hl;
                            grade[i].period5 = calc;
                            grade[i].period6 = coding;
                            }
                        }
                        if(fifthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(calc) || grade[i].preference2.equals(calc) || grade[i].preference3.equals(calc))){
                            if(sixthPeriod[l].equals(grade[i].preference1) && (grade[i].preference1.equals(rhetoric) || grade[i].preference2.equals(rhetoric) || grade[i].preference3.equals(rhetoric))){
                            grade[i].period1 = french;
                            grade[i].period2 = art;
                            grade[i].period3 = hl;
                            grade[i].period4 = hl;
                            grade[i].period5 = calc;
                            grade[i].period6 = rhetoric;
                            }
                        }
                        if(fifthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(coding) || grade[i].preference2.equals(coding) || grade[i].preference3.equals(coding))){
                            if(sixthPeriod[l].equals(grade[i].preference1) && (grade[i].preference1.equals(calc) || grade[i].preference2.equals(calc) || grade[i].preference3.equals(calc))){
                            grade[i].period1 = french;
                            grade[i].period2 = art;
                            grade[i].period3 = hl;
                            grade[i].period4 = hl;
                            grade[i].period5 = coding;
                            grade[i].period6 = calc;
                            }
                        }
                        if(fifthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(coding) || grade[i].preference2.equals(coding) || grade[i].preference3.equals(coding))){
                            if(sixthPeriod[l].equals(grade[i].preference1) && (grade[i].preference1.equals(rhetoric) || grade[i].preference2.equals(rhetoric) || grade[i].preference3.equals(rhetoric))){
                            grade[i].period1 = french;
                            grade[i].period2 = art;
                            grade[i].period3 = hl;
                            grade[i].period4 = hl;
                            grade[i].period5 = coding;
                            grade[i].period6 = rhetoric;
                            }
                        }
                    }
                    }
                }
            }
            }
        }
        
        for(int j = 0; j < secondPeriod.length; j++){
            if(secondPeriod[j].equals(grade[i].preference1) && (grade[i].preference1.equals(french) || grade[i].preference2.equals(french) || grade[i].preference3.equals(french))){
                grade[i].period2 = french;
                grade[i].period1 = art;
            }
        }
        
        for(int j = 0; j < thirdPeriod.length; j++){
            if(thirdPeriod[j].equals(french)){
                grade[i].period3 = french;
                grade[i].period5 = hl;
                grade[i].period6 = hl;
            }
        }
        for(int j = 0; j < fourthPeriod.length; j++){
            if(fourthPeriod[j].equals(french)){
                grade[i].period4 = french;
                grade[i].period5 = hl;
                grade[i].period6 = hl;
            }
        }
        for(int j = 0; j < fifthPeriod.length; j++){
            if(fifthPeriod[j].equals(french)){
                grade[i].period5 = french;
                grade[i].period3 = hl;
                grade[i].period4 = hl;
            }
        }
        for(int j = 0; j < sixthPeriod.length; j++){
            if(sixthPeriod[j].equals(french)){
                grade[i].period6 = french;
                grade[i].period3 = hl;
                grade[i].period4 = hl;
            }
        }
                        
        for(int j = 0; j < firstPeriod.length; j++){
            if(firstPeriod[j].equals(grade[i].preference1) && (grade[i].preference1.equals(greek) || grade[i].preference2.equals(greek) || grade[i].preference3.equals(greek))){
            grade[i].period1 = greek;
            grade[i].period2 = art;
            for(int h = 0; h < thirdPeriod.length; h++){
                if(thirdPeriod[h].equals(grade[i].preference1) && (grade[i].preference1.equals(coding) || grade[i].preference2.equals(coding) || grade[i].preference3.equals(coding))){
                    grade[i].period3 = coding;
                    for(int k = 0; k < fourthPeriod.length; k++){
                        if(fourthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(rhetoric) || grade[i].preference2.equals(rhetoric) || grade[i].preference3.equals(rhetoric))){
                            grade[i].period1 = greek;
                            grade[i].period2 = art;
                            grade[i].period3 = coding;
                            grade[i].period4 = rhetoric;
                            grade[i].period5 = hl;
                            grade[i].period6 = hl;
                        }
                        if(fourthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(physics) || grade[i].preference2.equals(physics) || grade[i].preference3.equals(physics))){
                            grade[i].period1 = greek;
                            grade[i].period2 = art;
                            grade[i].period3 = coding;
                            grade[i].period4 = physics;
                            grade[i].period5 = hl;
                            grade[i].period6 = hl;
                        }
                        if(fourthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(calc) || grade[i].preference2.equals(calc) || grade[i].preference3.equals(calc))){
                            grade[i].period1 = greek;
                            grade[i].period2 = art;
                            grade[i].period3 = coding;
                            grade[i].period4 = calc;
                            grade[i].period5 = hl;
                            grade[i].period6 = hl;
                        }
                    }
                }
                
                if(thirdPeriod[h].equals(grade[i].preference1) && (grade[i].preference1.equals(calc) || grade[i].preference2.equals(calc) || grade[i].preference3.equals(calc))){
                    grade[i].period3 = calc;
                    for(int k = 0; k < fourthPeriod.length; k++){
                        if(fourthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(rhetoric) || grade[i].preference2.equals(rhetoric) || grade[i].preference3.equals(rhetoric))){
                            grade[i].period1 = greek;
                            grade[i].period2 = art;
                            grade[i].period3 = calc;
                            grade[i].period4 = rhetoric;
                            grade[i].period5 = hl;
                            grade[i].period6 = hl;
                        }
                        if(fourthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(physics) || grade[i].preference2.equals(physics) || grade[i].preference3.equals(physics))){
                            grade[i].period1 = greek;
                            grade[i].period2 = art;
                            grade[i].period3 = calc;
                            grade[i].period4 = physics;
                            grade[i].period5 = hl;
                            grade[i].period6 = hl;
                        }
                        if(fourthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(coding) || grade[i].preference2.equals(coding) || grade[i].preference3.equals(coding))){
                            grade[i].period1 = greek;
                            grade[i].period2 = art;
                            grade[i].period3 = calc;
                            grade[i].period4 = coding;
                            grade[i].period5 = hl;
                            grade[i].period6 = hl;
                        }
                    }
                }
                
                if(thirdPeriod[h].equals(grade[i].preference1) && (grade[i].preference1.equals(physics) || grade[i].preference2.equals(physics) || grade[i].preference3.equals(physics))){
                    grade[i].period3 = physics;
                    for(int k = 0; k < fourthPeriod.length; k++){
                        if(fourthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(rhetoric) || grade[i].preference2.equals(rhetoric) || grade[i].preference3.equals(rhetoric))){
                            grade[i].period1 = greek;
                            grade[i].period2 = art;
                            grade[i].period3 = physics;
                            grade[i].period4 = rhetoric;
                            grade[i].period5 = hl;
                            grade[i].period6 = hl;
                        }
                        if(fourthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(calc) || grade[i].preference2.equals(calc) || grade[i].preference3.equals(calc))){
                            grade[i].period1 = greek;
                            grade[i].period2 = art;
                            grade[i].period3 = physics;
                            grade[i].period4 = calc;
                            grade[i].period5 = hl;
                            grade[i].period6 = hl;
                        }
                        if(fourthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(coding) || grade[i].preference2.equals(coding) || grade[i].preference3.equals(coding))){
                            grade[i].period1 = greek;
                            grade[i].period2 = art;
                            grade[i].period3 = physics;
                            grade[i].period4 = coding;
                            grade[i].period5 = hl;
                            grade[i].period6 = hl;
                        }
                    }
                }
                
                if(thirdPeriod[h].equals(grade[i].preference1) && (grade[i].preference1.equals(hl) || grade[i].preference2.equals(hl) || grade[i].preference3.equals(hl))){
                    grade[i].period3 = hl;
                    grade[i].period4 = hl;
                    for(int k = 0; k < fifthPeriod.length; k++){
                        for(int l = 0; l < sixthPeriod.length; l++){
                        if(fifthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(rhetoric) || grade[i].preference2.equals(rhetoric) || grade[i].preference3.equals(rhetoric))){
                            if(sixthPeriod[l].equals(grade[i].preference1) && (grade[i].preference1.equals(calc) || grade[i].preference2.equals(calc) || grade[i].preference3.equals(calc))){
                            grade[i].period1 = greek;
                            grade[i].period2 = art;
                            grade[i].period3 = hl;
                            grade[i].period4 = hl;
                            grade[i].period5 = rhetoric;
                            grade[i].period6 = calc;
                            }
                        }
                        if(fifthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(calc) || grade[i].preference2.equals(calc) || grade[i].preference3.equals(calc))){
                            if(sixthPeriod[l].equals(grade[i].preference1) && (grade[i].preference1.equals(coding) || grade[i].preference2.equals(coding) || grade[i].preference3.equals(coding))){
                            grade[i].period1 = greek;
                            grade[i].period2 = art;
                            grade[i].period3 = hl;
                            grade[i].period4 = hl;
                            grade[i].period5 = calc;
                            grade[i].period6 = coding;
                            }
                        }
                        if(fifthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(calc) || grade[i].preference2.equals(calc) || grade[i].preference3.equals(calc))){
                            if(sixthPeriod[l].equals(grade[i].preference1) && (grade[i].preference1.equals(rhetoric) || grade[i].preference2.equals(rhetoric) || grade[i].preference3.equals(rhetoric))){
                            grade[i].period1 = greek;
                            grade[i].period2 = art;
                            grade[i].period3 = hl;
                            grade[i].period4 = hl;
                            grade[i].period5 = calc;
                            grade[i].period6 = rhetoric;
                            }
                        }
                        if(fifthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(coding) || grade[i].preference2.equals(coding) || grade[i].preference3.equals(coding))){
                            if(sixthPeriod[l].equals(grade[i].preference1) && (grade[i].preference1.equals(calc) || grade[i].preference2.equals(calc) || grade[i].preference3.equals(calc))){
                            grade[i].period1 = greek;
                            grade[i].period2 = art;
                            grade[i].period3 = hl;
                            grade[i].period4 = hl;
                            grade[i].period5 = coding;
                            grade[i].period6 = calc;
                            }
                        }
                        if(fifthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(coding) || grade[i].preference2.equals(coding) || grade[i].preference3.equals(coding))){
                            if(sixthPeriod[l].equals(grade[i].preference1) && (grade[i].preference1.equals(rhetoric) || grade[i].preference2.equals(rhetoric) || grade[i].preference3.equals(rhetoric))){
                            grade[i].period1 = greek;
                            grade[i].period2 = art;
                            grade[i].period3 = hl;
                            grade[i].period4 = hl;
                            grade[i].period5 = coding;
                            grade[i].period6 = rhetoric;
                            }
                        }
                    }
                    }
                }
            }
            }
        }
                            for(int j = 0; j < secondPeriod.length; j++){
                                if(secondPeriod[j].equals(greek)){
                                    grade[i].period2 = greek;
                                    grade[i].period1 = art;
                                }
                            }
                            for(int j = 0; j < thirdPeriod.length; j++){
                                if(thirdPeriod[j].equals(greek)){
                                    grade[i].period3 = greek;
                                    grade[i].period5 = hl;
                                    grade[i].period6 = hl;
                                }
                            }
                            for(int j = 0; j < fourthPeriod.length; j++){
                                if(fourthPeriod[j].equals(greek)){
                                    grade[i].period4 = greek;
                                    grade[i].period5 = hl;
                                    grade[i].period6 = hl;
                                }
                            }
                            for(int j = 0; j < fifthPeriod.length; j++){
                                if(fifthPeriod[j].equals(greek)){
                                    grade[i].period5 = greek;
                                    grade[i].period3 = hl;
                                    grade[i].period4 = hl;
                                }
                            }
                            for(int j = 0; j < sixthPeriod.length; j++){
                                if(sixthPeriod[j].equals(greek)){
                                    grade[i].period6 = greek;
                                    grade[i].period3 = hl;
                                    grade[i].period4 = hl;
                                }
                            }
    
        for(int j = 0; j < firstPeriod.length; j++){
            if(firstPeriod[j].equals(grade[i].preference1) && (grade[i].preference1.equals(german) || grade[i].preference2.equals(german) || grade[i].preference3.equals(german))){
            grade[i].period1 = german;
            grade[i].period2 = art;
            for(int h = 0; h < thirdPeriod.length; h++){
                if(thirdPeriod[h].equals(grade[i].preference1) && (grade[i].preference1.equals(coding) || grade[i].preference2.equals(coding) || grade[i].preference3.equals(coding))){
                    grade[i].period3 = coding;
                    for(int k = 0; k < fourthPeriod.length; k++){
                        if(fourthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(rhetoric) || grade[i].preference2.equals(rhetoric) || grade[i].preference3.equals(rhetoric))){
                            grade[i].period1 = german;
                            grade[i].period2 = art;
                            grade[i].period3 = coding;
                            grade[i].period4 = rhetoric;
                            grade[i].period5 = hl;
                            grade[i].period6 = hl;
                        }
                        if(fourthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(physics) || grade[i].preference2.equals(physics) || grade[i].preference3.equals(physics))){
                            grade[i].period1 = german;
                            grade[i].period2 = art;
                            grade[i].period3 = coding;
                            grade[i].period4 = physics;
                            grade[i].period5 = hl;
                            grade[i].period6 = hl;
                        }
                        if(fourthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(calc) || grade[i].preference2.equals(calc) || grade[i].preference3.equals(calc))){
                            grade[i].period1 = german;
                            grade[i].period2 = art;
                            grade[i].period3 = coding;
                            grade[i].period4 = calc;
                            grade[i].period5 = hl;
                            grade[i].period6 = hl;
                        }
                    }
                }
                
                if(thirdPeriod[h].equals(grade[i].preference1) && (grade[i].preference1.equals(calc) || grade[i].preference2.equals(calc) || grade[i].preference3.equals(calc))){
                    grade[i].period3 = calc;
                    for(int k = 0; k < fourthPeriod.length; k++){
                        if(fourthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(rhetoric) || grade[i].preference2.equals(rhetoric) || grade[i].preference3.equals(rhetoric))){
                            grade[i].period1 = german;
                            grade[i].period2 = art;
                            grade[i].period3 = calc;
                            grade[i].period4 = rhetoric;
                            grade[i].period5 = hl;
                            grade[i].period6 = hl;
                        }
                        if(fourthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(physics) || grade[i].preference2.equals(physics) || grade[i].preference3.equals(physics))){
                            grade[i].period1 = german;
                            grade[i].period2 = art;
                            grade[i].period3 = calc;
                            grade[i].period4 = physics;
                            grade[i].period5 = hl;
                            grade[i].period6 = hl;
                        }
                        if(fourthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(coding) || grade[i].preference2.equals(coding) || grade[i].preference3.equals(coding))){
                            grade[i].period1 = german;
                            grade[i].period2 = art;
                            grade[i].period3 = calc;
                            grade[i].period4 = coding;
                            grade[i].period5 = hl;
                            grade[i].period6 = hl;
                        }
                    }
                }
                
                if(thirdPeriod[h].equals(grade[i].preference1) && (grade[i].preference1.equals(physics) || grade[i].preference2.equals(physics) || grade[i].preference3.equals(physics))){
                    grade[i].period3 = physics;
                    for(int k = 0; k < fourthPeriod.length; k++){
                        if(fourthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(rhetoric) || grade[i].preference2.equals(rhetoric) || grade[i].preference3.equals(rhetoric))){
                            grade[i].period1 = german;
                            grade[i].period2 = art;
                            grade[i].period3 = physics;
                            grade[i].period4 = rhetoric;
                            grade[i].period5 = hl;
                            grade[i].period6 = hl;
                        }
                        if(fourthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(calc) || grade[i].preference2.equals(calc) || grade[i].preference3.equals(calc))){
                            grade[i].period1 = german;
                            grade[i].period2 = art;
                            grade[i].period3 = physics;
                            grade[i].period4 = calc;
                            grade[i].period5 = hl;
                            grade[i].period6 = hl;
                        }
                        if(fourthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(coding) || grade[i].preference2.equals(coding) || grade[i].preference3.equals(coding))){
                            grade[i].period1 = german;
                            grade[i].period2 = art;
                            grade[i].period3 = physics;
                            grade[i].period4 = coding;
                            grade[i].period5 = hl;
                            grade[i].period6 = hl;
                        }
                    }
                }
                
                if(thirdPeriod[h].equals(grade[i].preference1) && (grade[i].preference1.equals(hl) || grade[i].preference2.equals(hl) || grade[i].preference3.equals(hl))){
                    grade[i].period3 = hl;
                    grade[i].period4 = hl;
                    for(int k = 0; k < fifthPeriod.length; k++){
                        for(int l = 0; l < sixthPeriod.length; l++){
                        if(fifthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(rhetoric) || grade[i].preference2.equals(rhetoric) || grade[i].preference3.equals(rhetoric))){
                            if(sixthPeriod[l].equals(grade[i].preference1) && (grade[i].preference1.equals(calc) || grade[i].preference2.equals(calc) || grade[i].preference3.equals(calc))){
                            grade[i].period1 = german;
                            grade[i].period2 = art;
                            grade[i].period3 = hl;
                            grade[i].period4 = hl;
                            grade[i].period5 = rhetoric;
                            grade[i].period6 = calc;
                            }
                        }
                        if(fifthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(calc) || grade[i].preference2.equals(calc) || grade[i].preference3.equals(calc))){
                            if(sixthPeriod[l].equals(grade[i].preference1) && (grade[i].preference1.equals(coding) || grade[i].preference2.equals(coding) || grade[i].preference3.equals(coding))){
                            grade[i].period1 = german;
                            grade[i].period2 = art;
                            grade[i].period3 = hl;
                            grade[i].period4 = hl;
                            grade[i].period5 = calc;
                            grade[i].period6 = coding;
                            }
                        }
                        if(fifthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(calc) || grade[i].preference2.equals(calc) || grade[i].preference3.equals(calc))){
                            if(sixthPeriod[l].equals(grade[i].preference1) && (grade[i].preference1.equals(rhetoric) || grade[i].preference2.equals(rhetoric) || grade[i].preference3.equals(rhetoric))){
                            grade[i].period1 = german;
                            grade[i].period2 = art;
                            grade[i].period3 = hl;
                            grade[i].period4 = hl;
                            grade[i].period5 = calc;
                            grade[i].period6 = rhetoric;
                            }
                        }
                        if(fifthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(coding) || grade[i].preference2.equals(coding) || grade[i].preference3.equals(coding))){
                            if(sixthPeriod[l].equals(grade[i].preference1) && (grade[i].preference1.equals(calc) || grade[i].preference2.equals(calc) || grade[i].preference3.equals(calc))){
                            grade[i].period1 = german;
                            grade[i].period2 = art;
                            grade[i].period3 = hl;
                            grade[i].period4 = hl;
                            grade[i].period5 = coding;
                            grade[i].period6 = calc;
                            }
                        }
                        if(fifthPeriod[k].equals(grade[i].preference1) && (grade[i].preference1.equals(coding) || grade[i].preference2.equals(coding) || grade[i].preference3.equals(coding))){
                            if(sixthPeriod[l].equals(grade[i].preference1) && (grade[i].preference1.equals(rhetoric) || grade[i].preference2.equals(rhetoric) || grade[i].preference3.equals(rhetoric))){
                            grade[i].period1 = german;
                            grade[i].period2 = art;
                            grade[i].period3 = hl;
                            grade[i].period4 = hl;
                            grade[i].period5 = coding;
                            grade[i].period6 = rhetoric;
                            }
                        }
                    }
                    }
                }
            }
            }
        }
                            
                            for(int j = 0; j < secondPeriod.length; j++){
                                if(secondPeriod[j].equals(german)){
                                    grade[i].period2 = german;
                                    grade[i].period1 = art;
                                }
                            }
                            for(int j = 0; j < thirdPeriod.length; j++){
                                if(thirdPeriod[j].equals(german)){
                                    grade[i].period3 = german;
                                    grade[i].period5 = hl;
                                    grade[i].period6 = hl;
                                }
                            }
                            for(int j = 0; j < fourthPeriod.length; j++){
                                if(fourthPeriod[j].equals(german)){
                                    grade[i].period4 = german;
                                    grade[i].period5 = hl;
                                    grade[i].period6 = hl;
                                }
                            }
                            for(int j = 0; j < fifthPeriod.length; j++){
                                if(fifthPeriod[j].equals(german)){
                                    grade[i].period5 = german;
                                    grade[i].period3 = hl;
                                    grade[i].period4 = hl;
                                }
                            }
                            for(int j = 0; j < sixthPeriod.length; j++){
                                if(sixthPeriod[j].equals(german)){
                                    grade[i].period6 = german;
                                    grade[i].period3 = hl;
                                    grade[i].period4 = hl;
                                }
                            }
        for(int j = 0; j < firstPeriod.length; j++){
                                if(firstPeriod[j].equals(spanish)){
                                    grade[i].period1 = spanish;
                                    grade[i].period2 = art;
                                    if(firstPeriod[j].equals(coding)){
                                        grade[i].period1 = coding;
                                        grade[i].period2 = art;
                                    }
                                }
                            }
                            for(int j = 0; j < secondPeriod.length; j++){
                                if(secondPeriod[j].equals(spanish)){
                                    grade[i].period2 = spanish;
                                    grade[i].period1 = art;
                                }
                            }
                            for(int j = 0; j < thirdPeriod.length; j++){
                                if(thirdPeriod[j].equals(spanish)){
                                    grade[i].period3 = spanish;
                                    grade[i].period5 = hl;
                                    grade[i].period6 = hl;
                                }
                            }
                            for(int j = 0; j < fourthPeriod.length; j++){
                                if(fourthPeriod[j].equals(spanish)){
                                    grade[i].period4 = spanish;
                                    grade[i].period5 = hl;
                                    grade[i].period6 = hl;
                                }
                            }
                            for(int j = 0; j < fifthPeriod.length; j++){
                                if(fifthPeriod[j].equals(spanish)){
                                    grade[i].period5 = spanish;
                                    grade[i].period3 = hl;
                                    grade[i].period4 = hl;
                                }
                            }
                            for(int j = 0; j < sixthPeriod.length; j++){
                                if(sixthPeriod[j].equals(spanish)){
                                grade[i].period6 = spanish;
                                grade[i].period3 = hl;
                                grade[i].period4 = hl;
                                }
                            }
    }
    }
    catch (FileNotFoundException e) {
    e.printStackTrace();
    } catch (IOException e) {
    e.printStackTrace();
    }
}
public static void main (String [] args){
    try{
        String filename = "C:\\Users\\ishik\\Downloads\\SeniorSchedulingCopy.xlsx";
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Output Template");
        int rowTotal = sheet.getPhysicalNumberOfRows();
        
        Student period[] = new Student[rowTotal];
        
        XSSFRow rowhead = sheet.createRow((short)0);
        rowhead.createCell(0).setCellValue("First Name");
        rowhead.createCell(1).setCellValue("Last Name");
        rowhead.createCell(2).setCellValue("1st Hour");
        rowhead.createCell(3).setCellValue("2nd Hour");
        rowhead.createCell(4).setCellValue("3rd Hour");
        rowhead.createCell(5).setCellValue("4th Hour");
        rowhead.createCell(6).setCellValue("5th Hour");
        rowhead.createCell(7).setCellValue("6th Hour");
        
        XSSFRow row = sheet.createRow((short)rowTotal);
        for(int i=1;i<rowTotal;i++){ 
            XSSFRow row1 = sheet.getRow(i);
            XSSFCell cellA1 = row1.getCell((short) 0);
            String period1 = (String)cellA1.getStringCellValue();
            XSSFCell cellB1 = row1.getCell((short) 1);
            String period2 = (String)cellB1.getStringCellValue();
            XSSFCell cellC1 = row1.getCell((short) 2);
            String period3 = (String)cellC1.getStringCellValue();
            XSSFCell cellD1 = row1.getCell((short) 3);
            String period4 = (String)cellD1.getStringCellValue();
            XSSFCell cellE1 = row1.getCell((short) 4);
            String period5 = cellE1.getStringCellValue();
            XSSFCell cellF1 = row1.getCell((short) 5);
            String period6 = cellF1.getStringCellValue();
            
            period[i-1] = new Student();
            period[i-1].period1 = period1;
            period[i-1].period2 = period2;
            period[i-1].period3 = period3;
            period[i-1].period4 = period4;
            period[i-1].period5 = period5;
            period[i-1].period6 = period6;
        }
        FileOutputStream fileOut = new FileOutputStream(filename);
        workbook.write(fileOut);
        fileOut.close();
        workbook.close();
        System.out.println("Excel file has been generated successfully.");
    }
    catch (Exception e){
    e.printStackTrace();
    }
}
}
*/

