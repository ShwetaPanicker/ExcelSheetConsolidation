

package excelLarge;
import com.monitorjbl.xlsx.StreamingReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.util.ArrayList;
import org.apache.poi.hssf.record.CFRuleBase.ComparisonOperator;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FontFormatting;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PatternFormatting;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFConditionalFormattingRule;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFFontFormatting;
import org.apache.poi.xssf.usermodel.XSSFPatternFormatting;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSheetConditionalFormatting;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.xml.sax.SAXException;


public class Consolidation
{ 
    
    
  File ExcelFileToRead;
  int rowStart,rowEnd,cn, d,a,lastColumn;
  ArrayList<String> bookList = new ArrayList<>();
  String pathOfNewFile;
  String pathOfDataFiles;
  String sheetname, str1, book;
  Double ans;
  XSSFCellStyle style;
  XSSFRow currentRow;
  Cell currentCell;
  XSSFCell cell;
  Workbook workbook1;
  Workbook workbook;
  Sheet datatypeSheet,newsheet;
  Cell newcell;
  XSSFCell cA,cA1;
  Row newrow, sheetrow;
  CellStyle newCellStyle ;
  XSSFFont font;
  
  XSSFRichTextString rts;
  XSSFSheetConditionalFormatting my_cond_format_layer ;
  XSSFConditionalFormattingRule my_rule;
  XSSFFontFormatting my_rule_pattern;
  XSSFPatternFormatting fill_pattern;
  
 public boolean readWriteXLSXFile(String src ,String des) throws IOException, InvalidFormatException, OpenXML4JException, SAXException
 {
   
  pathOfNewFile = des;
  pathOfDataFiles=src;
  workbook1= new XSSFWorkbook();
  
  File file= new File(src);
  File[] files = file.listFiles();
   for(File f: files)
   {  
    if(f.getName().endsWith(".xlsx"))
    bookList.add(f.getName());
   }

   for (String b : bookList)
   {
     book=b;  
    ///ExcelFileToRead = new FileInputStream(pathOfDataFiles+b);
    ExcelFileToRead=new File(pathOfDataFiles+book);
    sheetname=book.replace(".xlsx",""); 
    //System.out.println(sheetname);
    workbook = StreamingReader.builder()
                        .rowCacheSize(200)    // number of rows to keep in memory (defaults to 10)
                        .bufferSize(4024) // buffer size to use when reading InputStream to file (defaults to 1024)
                        //.sheetIndex// index of sheet to use (defaults to 0)
                          // name of sheet to use (overrides sheetIndex)
                        .open(ExcelFileToRead); 
    
    datatypeSheet = workbook.getSheetAt(0);
    
    newsheet = workbook1.createSheet(sheetname);
    
    //rowIndex=0;
    
    if(datatypeSheet==null)
           System.out.println("NULL");
    
     for (Row r: datatypeSheet)
    {
      int rowCount=r.getRowNum();
      newrow=newsheet.createRow(rowCount);
     // cellIndex=0;
      if (r == null)
      {
      
      }

      else
      {
       //lastColumn= r.getLastCellNum();
       for (Cell c: r)
       {

        currentCell= c;
        if(currentCell==null)
        {
         //cellIndex++;
        }
        else
        {
           switch (currentCell.getCellType())
          {
               
           case XSSFCell.CELL_TYPE_STRING:
            int cellCount= currentCell.getColumnIndex();
            newcell=newrow.createCell(cellCount);
            newcell.setCellValue(currentCell.getStringCellValue());
//            newCellStyle = newcell.getSheet().getWorkbook().createCellStyle();
//            newCellStyle.cloneStyleFrom(currentCell.getCellStyle());
//            newcell.setCellStyle(newCellStyle);
            //cellIndex++;
           break;

           case XSSFCell.CELL_TYPE_NUMERIC:
            cellCount= currentCell.getColumnIndex();
            newcell=newrow.createCell(cellCount);
            newcell.setCellValue(currentCell.getNumericCellValue());
//            newCellStyle = newcell.getSheet().getWorkbook().createCellStyle();
//            newCellStyle.cloneStyleFrom(currentCell.getCellStyle());
//            newcell.setCellStyle(newCellStyle);
            //cellIndex++;          
           break;

           default:
           break;
          }
        }

       }  
      }
       
    }
   newsheet.autoSizeColumn(0);
   newsheet.autoSizeColumn(1);
   newsheet.autoSizeColumn(2);
   newsheet.autoSizeColumn(3);
   newsheet.autoSizeColumn(4);
   
  
   System.out.println("\nThe first sheet from file "+b+ " written successfully in new excel file");

  }           
         try (FileOutputStream fileOut = new FileOutputStream(pathOfNewFile))
  {
   workbook1.write(fileOut);
   fileOut.flush();
   fileOut.close();
   workbook1.close();
   }
  
  FormatFile();
  return true;

}
 
public void FormatFile() throws FileNotFoundException, IOException
{
    XSSFWorkbook fworkbook;
      try (InputStream ExcelFileToFormat = new FileInputStream(pathOfNewFile))
      {
          fworkbook = new XSSFWorkbook(ExcelFileToFormat);
          XSSFSheet fSheet;
          int n=fworkbook.getNumberOfSheets();
          
          for(int j=0; j<n;j++)
          {
              fSheet = fworkbook.getSheetAt(j);
              a=fSheet.getLastRowNum();
              
              DataFormatter df = new DataFormatter();
              cA = fSheet.getRow(4).getCell(2);
              
              String denominator = df.formatCellValue(cA);
              
             
              
              for(int i=17; i<a; i++)
              {
                  
                  //try
                  
                  cA1 = fSheet.getRow(i).getCell(2);
                  
                  
                  //catch(Exception e)
                  
                  
                  String numerator = df.formatCellValue(cA1);
                  
                  if(numerator  == null)
                  {
                      System.out.println("NULL");
                      break;
                  }
                  if(denominator  == null)
                  {
                      System.out.println("NULL");
                      break;
                  }
                  
                  
                  
                  
                  //DecimalFormat df1 = new DecimalFormat("#.##%");
                  //DecimalFormat df2 = new DecimalFormat("#.##");
                  
                  ans = (Double.parseDouble(numerator)/Double.parseDouble(denominator));
                  
                  style = fworkbook.createCellStyle();
                  style.setDataFormat(fworkbook.createDataFormat().getFormat("0.00%"));
                  cell=fSheet.getRow(i).createCell(4);
                  cell.setCellValue(ans);
                  cell.setCellStyle(style);
                  my_cond_format_layer = fSheet.getSheetConditionalFormatting();
                  my_rule = my_cond_format_layer.createConditionalFormattingRule(ComparisonOperator.GT, "20%");
                  my_rule_pattern = my_rule.createFontFormatting();
                  my_rule_pattern.setFontColorIndex(IndexedColors.RED.getIndex());
                  CellRangeAddress[] my_data_range = {CellRangeAddress.valueOf("E17:E9999")};
                  my_cond_format_layer.addConditionalFormatting(my_data_range,my_rule);
                  fill_pattern = my_rule.createPatternFormatting();
                  fill_pattern.setFillBackgroundColor(IndexedColors.YELLOW.index);
                  
                  
              }
              
              XSSFRow fsheetrow = fSheet.getRow(15);
              if(fsheetrow == null)
              {
                  fsheetrow = fSheet.createRow(14);
              }
              
              //Update the value of cell
              
              
              newcell = fsheetrow.getCell(4);
              if(newcell == null)
              {
                  newcell = fsheetrow.createCell(4);
                  
              }
              
              str1="#DIFFERENCES IN PERCENTAGE";
              
              rts=new XSSFRichTextString(str1);
              
              font=fworkbook.createFont();
              font.setBold(true);
              font.setUnderline((byte)1);
              rts.applyFont(font);
              newcell.setCellValue(rts);
              fSheet.autoSizeColumn(0);
              fSheet.autoSizeColumn(1);
              fSheet.autoSizeColumn(2);
              fSheet.autoSizeColumn(3);
              fSheet.autoSizeColumn(4);
              
          }    }

    FileOutputStream outStream = new FileOutputStream(pathOfNewFile);
                    
        fworkbook.write(outStream);
        outStream.flush();
        outStream.close();
        fworkbook.close();

}
     
 
 
 


 public static void main(String[] args) throws Exception
 {
 // TODO code application logic here

    String src="C:\\Users\\Shweta\\Documents\\EXPERIMENT\\";
    String des="C:\\Users\\Shweta\\Documents\\EXPERIMENT\\NewFileFolder\\File.xlsx";
    ExcelLarge read= new ExcelLarge();
  if(read.readWriteXLSXFile(src, des))
  {
  System.out.println("\nYeyy Done \nExcel Read and Write Operation from all files in the folder with CellStyle  Successful :)\n");
  }
  else
  {
  System.out.println("\nFailed :( ");
  }      
 }

}
