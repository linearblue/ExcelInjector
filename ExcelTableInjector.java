package com.linearblue.exceltableinjector;

/*
 * 
 */

import java.io.*;
import java.io.FileOutputStream;
import java.net.URI;
import java.net.URL;
import java.util.Iterator;
import java.util.List;
import java.util.HashMap;

import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.DateUtil;

import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumn;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumns;

import org.w3c.dom.*;

/**
 * @author jasonerickson
 */
public class ExcelTableInjector {
    
    static final String multipleUsage = "Usage: -c injectMultiple -s sourcePath.xlsx -d destPath.xlsx -n sheetNumber [-p dataFilePath] -b table:table1=dataPath1~table:tableN=dataPathN~cell:cellPosition1=value1~cell:cellPositionN=valueN ... any combo of table and cell entries separted by ~";
    static final String cellUsage = "Usage: -s sourcePath.xlsx -d destPath.xlsx -n sheetNumber -a cellPosition -v value";
    static final String injectUsage = "Usage: -s sourcePath.xlsx -d destPath.xlsx -n sheetNumber -t tableName -m dataPath.mer";
    static final boolean mac = System.getProperty("os.name").startsWith("Mac");
    
    HashMap<Integer,String> colNamesHash;
    FileInputStream input_document = null;
    XSSFWorkbook workbook = null;
    XSSFSheet worksheet = null;
    List<XSSFTable> list = null;
    CTTable cttable = null;
    XSSFTable xtable = null;
    HashMap<Integer,HashMap> rowHash;
    int sheetNum = 0;
    
    public ExcelTableInjector() throws Exception {
        
    }
    
    public class CellDef {
        
        public int type;
        public int index;
        public String xName;
        XSSFCellStyle style;
        
        public CellDef(int type, int index, String xName, XSSFCellStyle style) {
            this.type = type;
            this.index = index;
            this.xName = xName;
            this.style = style;
        }
       
    }
    
    private String readFile(File source) throws Exception {
        StringBuilder buffer = new StringBuilder();
        String encoding = "UTF-8";
        Reader in = null;
        try {
            FileInputStream fis = new FileInputStream(source);
            InputStreamReader isr = new InputStreamReader(fis, encoding);
            in = new BufferedReader(isr);
            int ch;
            while ((ch = in.read()) > -1) {
                buffer.append((char) ch);
            }
            return buffer.toString();
        } finally {
            if (in != null) {
                in.close();
            }
        }
    }
     
    private void loadData(String dataPath) throws Exception {
         // build data
        File dataFile = new File(dataPath);
        String data = null;
        try {
            data = readFile(dataFile);
        } catch (Exception e) {
            throw new Exception("Unable to read file: " + dataFile.getAbsolutePath());
        }
        if(data == null || data.length()==0)
            throw new Exception("Data file empty or invalid.");
        
        String[] dataArray = null;
        try {
             dataArray = mac ? data.split("\r") : data.split("\n");
        } catch (Exception e) {
            throw new Exception("Data file not separated by return character.");
        }
        if(dataArray.length < 1) throw new Exception("Data file not separated by return character.");
        
        rowHash = new HashMap();
        this.colNamesHash = new HashMap();
        int rowCount = dataArray.length;
        for(int i=0; i<dataArray.length; i++) {
            String row = dataArray[i];
            if(i==0) {
                if(row.indexOf("\"")>-1) throw new Exception("Missing header row with field names.");
                String[] colNames = row.split(",");
                for(int j=0; j<colNames.length; j++) {
                    colNamesHash.put(j, colNames[j]);
                }
            } else {
                HashMap<Integer,String> dataColHash = new HashMap();
                if(row.indexOf("\"")<0) throw new Exception("Improper data format. Should be Merge, which uses quotes around values. [" + data + "] at path: [" + dataFile + "]");
                String[] colVals = row.split("\",\"");
                for(int j=0; j<colVals.length; j++) {
                    String val = colVals[j];
                    if(j==0) val = val.substring(1);
                    else if(j==(colVals.length-1)) val = val.substring(0,val.length()-1);
                    dataColHash.put(j, val);
                }
                rowHash.put(i,dataColHash);
            }
        }
        if(rowHash.size()<2) throw new Exception("At least two data rows are required.");
     }
     
    private void loadExcel(String xslSourcePath, int sheetNum) throws Exception {
        try {
            //Read Excel document first
            input_document = new FileInputStream(new File(xslSourcePath));
            // convert it into a POI object
            workbook = new XSSFWorkbook(input_document);
            // Read excel sheet that needs to be updated
        } catch (Exception e) {
            throw new Exception("Unable to load Excel file", e);
        }
     }
    
    private void loadTable(String tableName) throws Exception {
        try {
            String sheetNumTxt = tableName.indexOf("[")>-1 ? tableName.substring(tableName.indexOf("[")) : null;
            if(sheetNumTxt != null) {
                this.sheetNum = new Integer(sheetNumTxt.substring(1,sheetNumTxt.length()-1));
                tableName = tableName.substring(0, tableName.indexOf("["));
            } else {
                this.sheetNum = 0;
            }
            worksheet = workbook.getSheetAt(this.sheetNum);
            
            // Get tables and iterate to specified
            list = (List<XSSFTable>)worksheet.getTables();
            Iterator<XSSFTable> itor = list.iterator();
            cttable = null;
            xtable = null;
            while(itor.hasNext()) {
                XSSFTable table = itor.next();
                if(table.getName().equals(tableName)) {
                    xtable = table;
                    cttable = table.getCTTable();
                    break;
                }
            }
            if(xtable == null || cttable == null) throw new Exception ("table not found.");
        } catch (Exception e) {
            throw new Exception("Unable to load specified table " + tableName, e);
        }
    }
    
    private void saveExcel(String xslDestPath) throws Exception {
        try {
            // Close the document
            input_document.close();
        } catch (Exception e) {
            throw new Exception("Unable to close source Excel file", e);
        }
        
        try {
            // turn on calc auto refresh on open
            worksheet.setForceFormulaRecalculation(true);
            //Open FileOutputStream to write updates
            FileOutputStream output_file = new FileOutputStream(new File(xslDestPath));
            //write changes
            workbook.write(output_file);
            //close the stream
            output_file.close(); 
        } catch (Exception e) {
            throw new Exception("Unable to save new Excel file", e);
        }
    }
    
    private void shiftRange(int numRowsDown, int rowStart, int rowEnd, int colStart, int colEnd) throws Exception {
        // loop through from bottom row up to either SHIFT the whole row, or copy necessary columns, insert  new row below, insert the data from the cell and then delete all style and format from the original cell
        for(int j=rowEnd;j>rowStart;j--) {
            XSSFRow row = worksheet.getRow(j);
            XSSFRow newRow = worksheet.getRow(j+numRowsDown);
            if(row == null && newRow == null) 
                continue;
            else if (row == null) {
                worksheet.removeRow(newRow);
                continue;
            }
            if(newRow == null) {
                newRow = worksheet.createRow(j+numRowsDown);
            }
//            newRow = worksheet.getRow(j+numRowsDown);
            //if(row ==null) continue;
            for(int k=colStart;k<=colEnd;k++) {
                XSSFCell cell = row.getCell(k);
                if(cell != null) {
                    XSSFCell newCell = newRow.getCell(k)==null ? newRow.createCell(k) : newRow.getCell(k);
                    if(cell==null) continue;
                    if(cell.getCellType() == XSSFCell.CELL_TYPE_FORMULA) {
                        String calc = cell.getCellFormula();
                        newCell.setCellStyle(cell.getCellStyle());
                        cell.setCellType(XSSFCell.CELL_TYPE_BLANK);
                        this.setCellFormula(newCell, calc);
                    } else if(cell.getCellType()==XSSFCell.CELL_TYPE_BLANK) {
                        newCell.setCellType(XSSFCell.CELL_TYPE_BLANK);
                    } else {
                        newCell.copyCellFrom(cell, new CellCopyPolicy());
                        newCell.setCellStyle(cell.getCellStyle());
                    }
                }
            }
        }
    }
    
    private int checkForContent(int rowStart, int rowEnd, int colStart, int colEnd) throws Exception {
        for(int j=rowEnd;j>rowStart;j--) {
            XSSFRow row = worksheet.getRow(j);
            if(row ==null) continue;
            for(int k=colStart;k<=colEnd;k++) {
                XSSFCell cell = row.getCell(k);
                if(cell==null) continue;
                if(cell.getCellType() == XSSFCell.CELL_TYPE_BLANK) continue;
                return j;
            }
        }
        return 0;
    }
    
    private void makeRoomForData() throws Exception {
        try {
            // parse the current AreaReference
            String ref = cttable.getRef();
            if(ref == null) throw new Exception("Missing valid AreaReference");
            String[] rc = ref.split(":");
            String colLetter = rc[1].replaceAll("[^a-zA-Z]", "");
            String colLetterStart = rc[0].replaceAll("[^a-zA-Z]", "");
            
            // get needed row positions
            int tableStartRow = xtable.getStartCellReference().getRow();
            int newLastRow = tableStartRow + rowHash.size() + 1;
            int sheetLastRow = worksheet.getLastRowNum();
            
            // Check for a/some Totals Row(s)
            int totalsRowCount = (int)cttable.getTotalsRowCount();
            if(totalsRowCount>0) {
                newLastRow = newLastRow + totalsRowCount;
            }
            
            // get actual last data row
            int tableLastRow = xtable.getEndCellReference().getRow()-totalsRowCount;
            
            // test if table will need to be expanded
            if(newLastRow-tableStartRow > 2) {
                String newRef = rc[0] + ":" + colLetter + newLastRow;
                
                // get column bounds
                int colStart = xtable.getStartCellReference().getCol();
                int colEnd = xtable.getEndCellReference().getCol();
                
                // unless table hits bottom bounds of sheet, make room
                if(tableLastRow < sheetLastRow) {
                    // find out whether anything directly below table by walking each cell between current table end cell and future end cell and looking for any existing content
                    int shiftRowStart = this.checkForContent(tableLastRow, newLastRow, colStart, colEnd);
                    // ... and if so ...
                    if(shiftRowStart > 0) {
                        // move the defined range of cells down
                        this.shiftRange(newLastRow-tableLastRow-1-totalsRowCount, tableLastRow, sheetLastRow, colStart, colEnd);
                    }
                } else if(sheetLastRow < newLastRow) {
                    
                }
                // Reset the Table Range Reference
                cttable.setRef(newRef);
            }
        } catch (Exception e) {
            throw new Exception("Unable to reset size of table.", e);
        }
        
    }
    
    private void setCellFormula(XSSFCell cell, String calc) throws Exception {
        Element b = (Element) cell.getCTCell().getDomNode();
        Element f = b.getOwnerDocument().createElementNS("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "f");
        if(b.hasAttribute("t")) b.removeAttribute("t");
        if(b.hasChildNodes()) b.removeChild(b.getElementsByTagName("v").item(0));
        f.appendChild(b.getOwnerDocument().createTextNode(calc));
        b.appendChild(f);
    }
    
    private void populateRow(XSSFRow xrow, HashMap<Integer,String> dataRow, HashMap<Integer,String> colNamesHash, HashMap<String, CellDef> xColNameMap) throws Exception {
        int colCount = colNamesHash.size();
        for(int i=0; i<colCount; i++) {
            String fmColName = colNamesHash.get(i);
            //String xColName = map.get(fmColName);
            CellDef def = xColNameMap.get(fmColName);
            if(def != null) {
                String xColName = def.xName;
                String fmData = dataRow.get(i);
                int xColInt = def.index;
                XSSFCell cell = null;
                try {
                    cell = xrow.getCell(xColInt);
                } catch (Exception noCell) {
                    throw new Exception ("No cell found at " + xColInt + " on row: " + xrow + " for column named: " + xColName);
                }
                if(cell == null) cell = xrow.createCell(xColInt);
                cell.setCellStyle(def.style);
                int cellType = def.type;
                if (cellType == XSSFCell.CELL_TYPE_BLANK) {
                    cell.setCellType(XSSFCell.CELL_TYPE_BLANK);
                    //cell.setCellValue(fmData);
                } else if (cellType == XSSFCell.CELL_TYPE_BOOLEAN) {
                    cell.setCellValue(fmData != null && (fmData.compareToIgnoreCase("true")==0 || fmData.equals("1")) ? true : false);
                } else if (cellType == XSSFCell.CELL_TYPE_FORMULA) {
                    //cell.setCellFormula(fmData);
                    this.setCellFormula(cell, fmData);
                } else if (cellType == 10001) {
                     java.util.Date d = new java.util.Date(fmData);
                        cell.setCellValue(d);
                } else if (cellType == XSSFCell.CELL_TYPE_NUMERIC) {
                    cell.setCellValue(new Double(fmData));
                } else if (cellType == XSSFCell.CELL_TYPE_STRING) {
                    cell.setCellValue(fmData);
                } else if(DateUtil.isCellDateFormatted(cell)) {
                    java.util.Date d = new java.util.Date(fmData);
                    cell.setCellValue(d);
                } else {
                    cell.setCellValue(fmData);
                }
            }
        }
    }
    
    private void injectData() throws Exception {
        try {
            // build table map
            // Get the table's sheet
            XSSFSheet xsheet = xtable.getXSSFSheet();
            CellReference startCell = xtable.getStartCellReference();
            CellReference endCell = xtable.getEndCellReference();
            // Populate row
            int startRow = startCell.getRow();
            XSSFRow firstRow = xsheet.getRow(startRow+1);
            XSSFRow secondRow = xsheet.getRow(startRow+2);

            CTTableColumns tableColumns = cttable.getTableColumns();
            CTTableColumn[] cols = tableColumns.getTableColumnArray();
            //List <CTTableColumn> cols = tableColumns.getTableColumnList();
            HashMap<String,CellDef> xColsMap = new HashMap();
            for(int m=0; m<cols.length; m++) {
                CTTableColumn col = cols[m];
                String colName = col.getName();
                XSSFCell cell1 = firstRow.getCell(m+startCell.getCol());
                XSSFCell cell2 = secondRow.getCell(m+startCell.getCol());
                String fmName = null;
                try {
                    fmName = cell2.getStringCellValue();
                } catch (Exception e2) {
                    //throw new Exception ("Cell:" + m+startCell.getCol());
                    fmName = cell2.getRawValue();
                }
//                String fmName = cell2.getRawValue();
                
                int cellType = cell1.getCellType();
                XSSFCellStyle style = cell1.getCellStyle();
                int cellPos = cell1.getColumnIndex();
                try {
                    if(DateUtil.isCellDateFormatted(cell1)) cellType = 10001;
                } catch (Exception e) {
                    // ignore
                }
                xColsMap.put(fmName, new CellDef(cellType, cellPos, colName, style));
            }
            // Populate additional rows
            int rowCount = rowHash.size()+1;
            for(int l=1; l<rowCount; l++) {
                int thisRow = startRow + l;
                HashMap row = rowHash.get(l);
                XSSFRow xrow = xsheet.getRow(thisRow);
                if(xrow == null) {
                    xrow = xsheet.createRow(thisRow);
//                    xrow = xsheet.getRow(thisRow);
                }
                this.populateRow(xrow, rowHash.get(l), colNamesHash, xColsMap);
            }
        } catch (Exception e) {
            throw new Exception("Unable to parse table", e);
        }
    }
    
    public void excelTableInject (String xslSourcePath, String xslDestPath, int sheetNum, String tableName, String dataPath) throws Exception {
        // load CSV data
        this.loadData(dataPath);
        // load Excel file,
        this.loadExcel(xslSourcePath, sheetNum);
        //  get worksheet and table 
        this.loadTable(tableName);
        // move existing contents below table down and reset table range reference
        this.makeRoomForData();
        // push CSV data into table
        this.injectData();
        // close source and save destination
        this.saveExcel(xslDestPath);
    }
    
    public void excelTableInject (String tableName, String dataPath) throws Exception {
        // load CSV data
        this.loadData(dataPath);
        //  get worksheet and table 
        this.loadTable(tableName);
        // move existing contents below table down and reset table range reference
        this.makeRoomForData();
        // push CSV data into table
        this.injectData();
    }
    
    public void updateCellValue(String cellPosition, String value) throws Exception {
        
        String sheetNumTxt = cellPosition.indexOf("[")>-1 ? cellPosition.substring(cellPosition.indexOf("[")) : null;
            if(sheetNumTxt != null) {
                this.sheetNum = new Integer(sheetNumTxt.substring(0,sheetNumTxt.length()-1));
                cellPosition = cellPosition.substring(cellPosition.indexOf("["));
            } else {
                this.sheetNum = 0;
            }
            worksheet = workbook.getSheetAt(this.sheetNum);
        
        CellReference c = new CellReference(cellPosition);
        XSSFCell cell = worksheet.getRow(c.getRow()).getCell(c.getCol());
        if(cell == null) throw new Exception("Invalid cell reference:" + cellPosition);
        if(value == null) {
           cell.setCellType(XSSFCell.CELL_TYPE_BLANK);
        } else if(cell.getCellType()==XSSFCell.CELL_TYPE_FORMULA) {
            this.setCellFormula(cell, value);
        } else {
            cell.setCellValue(value);
        }
    }
    
    public void updateCellValue(String xslSourcePath, String xslDestPath, int sheetNum, String cellPosition, String value) throws Exception {
        // load Excel file,
        this.loadExcel(xslSourcePath, sheetNum);
        CellReference c = new CellReference(cellPosition);
        XSSFCell cell = worksheet.getRow(c.getRow()).getCell(c.getCol());
        if(cell == null) throw new Exception("Invalid cell reference:" + cellPosition);
        if(value == null) {
           cell.setCellType(XSSFCell.CELL_TYPE_BLANK);
        } else if(cell.getCellType()==XSSFCell.CELL_TYPE_FORMULA) {
            this.setCellFormula(cell, value);
        } else {
            cell.setCellValue(value);
        }
        // close source and save destination
        this.saveExcel(xslDestPath);
    }
    
    private void doMultiple(String batchInstructions, String dataPath, String sourcePath, String destPath, int sheet) throws Exception {
        this.loadExcel(sourcePath, sheet);
        String[] pairs = batchInstructions.split("~");
        for(int i=0; i<pairs.length; i++) {
            String[] pair = pairs[i].split("=");
            if(pair.length!=2) {
                System.err.println(multipleUsage);
                System.exit(-1);
            }
            String a = pair[0];
            String b = pair[1];
            if(a.startsWith("table:")) {
                if(dataPath != null) b = new File(dataPath, b).getAbsolutePath();
                this.excelTableInject(a.substring("table:".length()), b);
            } else if(a.startsWith("cell:")) {
                this.updateCellValue(a.substring("cell:".length()), b);
            } else {
                System.err.println(multipleUsage);
                System.exit(-1);
            }
        }
        this.saveExcel(destPath);
    }
    
    public static File getLocationFile() throws Exception {
        URL url = ExcelTableInjector.class.getProtectionDomain().getCodeSource().getLocation();
        URI uri = new URI(url.toString());
        String filePath = uri.getPath();
        filePath = filePath.substring(filePath.indexOf("/"), filePath.lastIndexOf("/"));
        File file = new File(filePath);
        return file;
    }
    
    private static void getStackTrace(Throwable thrown, StringBuilder buf) throws Exception {
        String lineReturn =  mac ? "\r" : "\n";
        StackTraceElement stack[] = thrown.getStackTrace();
        for (int index = 0; index < stack.length; index++) {
            String cname = stack[index].getClassName();
            if (cname != null) {
                buf.append("\t");
                buf.append(cname);
                buf.append(" - ");
            }
            String fname = stack[index].getMethodName();
            if (fname != null) {
                buf.append(fname);
                buf.append(" - ");
            }
            int line = stack[index].getLineNumber();
            if (line >= 0) {
                buf.append(line);
                buf.append(lineReturn);
            }
        }
    }
    
   private static String getStackTrace(Throwable thrown) {
        try {
            String lineReturn =  mac ? "\r" : "\n";
            StringBuilder buf = new StringBuilder();
            while (thrown != null) {
                String message = thrown.getMessage();
                if (message != null) {
                    buf.append(message);
                    buf.append(lineReturn);
                } else {
                    buf.append("[no message]");
                    buf.append(lineReturn);
                }
                getStackTrace(thrown, buf);
                thrown = thrown.getCause();
            }
            return buf.toString();
        } catch (Exception e2) {
            return "";
        }
    }
   
   private static void printStackTrace(Throwable thrown) {
       try {
           String content = getStackTrace(thrown);
           File root = ExcelTableInjector.getLocationFile();
           File log = new File(root, "ExcelInjector.log");
           String output = log.getAbsolutePath();
           if (output == null || content == null) {
                throw new Exception("Invalid file reference or content provided");
            }
            String encoding = "UTF-8";
            FileOutputStream outputStream = null;
            try {
                //File outputDir = output.getParentFile();
                //outputDir.mkdirs();
                outputStream = new FileOutputStream(output);
                OutputStreamWriter writer = null;
                PrintWriter printer = null;
                try {
                    writer = new OutputStreamWriter(outputStream, encoding);
                    printer = new PrintWriter(writer);
                    printer.print(content);
                } finally {
                    if (printer != null) {
                        printer.close();
                    }
                    if (writer != null) {
                        writer.close();
                    }
                }
            } finally {
                if (outputStream != null) {
                    outputStream.close();
                }
            }
       } catch (Exception e) {
           e.printStackTrace();
       }
   }
    
    public static void main(String[] args) {
        try {
            String sourcePath = null, destPath = null, dataPath = null, tableName = null, command = null, areaRef = null, value = null, batchInstructions = null, dataFolder = null;
            int sheet = 0;
            int paramNum = 0;
            StringBuffer buf = new StringBuffer();
            for (int i = 0; i < args.length; i++) {
                buf.append(args[i] + " ");
                if(args[i].startsWith("-")) {
                    if(args[i].compareToIgnoreCase("-s")==0) paramNum = 1;
                    else if(args[i].compareToIgnoreCase("-d")==0) paramNum = 2;
                    else if(args[i].compareToIgnoreCase("-m")==0) paramNum = 3;
                    else if(args[i].compareToIgnoreCase("-t")==0) paramNum = 4;
                    else if(args[i].compareToIgnoreCase("-n")==0) paramNum = 5;
                    else if(args[i].compareToIgnoreCase("-c")==0) paramNum = 6;
                    else if(args[i].compareToIgnoreCase("-a")==0) paramNum = 7;
                    else if(args[i].compareToIgnoreCase("-v")==0) paramNum = 8;
                    else if(args[i].compareToIgnoreCase("-b")==0) paramNum = 9;
                    else if(args[i].compareToIgnoreCase("-p")==0) paramNum = 10;
                    
                    
                    else paramNum = 0;
                } else {
                    if(paramNum ==0) continue;
                    else if(paramNum==1) sourcePath = (sourcePath != null ? sourcePath + " " : "") + args[i];
                    else if(paramNum==2) destPath = (destPath != null ? destPath + " " : "") + args[i];
                    else if(paramNum==3) dataPath = (dataPath != null ? dataPath + " " : "") + args[i];
                    else if(paramNum==4) tableName = (tableName != null ? tableName + " " : "") + args[i];
                    else if(paramNum==5) sheet = new Integer(args[i]);
                    else if(paramNum==6) command = (command != null ? command + " " : "") + args[i];
                    else if(paramNum==7) areaRef = (areaRef != null ? areaRef + " " : "") + args[i];
                    else if(paramNum==8) value = (value != null ? value + " " : "") + args[i];
                    else if(paramNum==9) batchInstructions = (batchInstructions != null ? batchInstructions + " " : "") + args[i];
                    else if(paramNum==10) dataFolder = (dataFolder != null ? dataFolder + " " : "") + args[i];
                }
            }
            
            File root = ExcelTableInjector.getLocationFile();
            
            boolean hasTableUpdateParams = (sourcePath == null || destPath == null || dataPath == null || tableName == null) ? false : true;
            boolean hasUpdateCellParams = (sourcePath == null || destPath == null || areaRef == null || value == null) ? false : true ;
            boolean hasBatchInstructions = (sourcePath == null || destPath == null || batchInstructions == null) ? false : true ;
            
            if(command != null) {
                if(command.compareToIgnoreCase("updateCell")==0 && hasUpdateCellParams == false) {
                    System.err.println(cellUsage);
                    System.exit(-1);
                } else if(command.compareToIgnoreCase("injectMultiple")==0 && hasBatchInstructions == false) {
                    System.err.println(multipleUsage);
                    System.exit(-1);
                }
            } else if(hasTableUpdateParams == false) {
                System.err.println(injectUsage);
                System.err.println("Received: " + buf.toString());
                System.exit(-1);
            }
            
            File sourceFile = null, destFile = null;
            
            if(dataFolder != null && sourcePath.indexOf("/")<0 && sourcePath.indexOf("\\")<0) {
                sourcePath = new File(dataFolder, sourcePath).getAbsolutePath();
                sourceFile = new File(sourcePath);
            } else {
                sourceFile = new File(sourcePath);
                if(!sourceFile.exists()) sourceFile = new File(root, sourcePath);
            }
            if(dataFolder != null && destPath.indexOf("/")<0 && destPath.indexOf("\\")<0) {
                destPath = new File(dataFolder, destPath).getAbsolutePath();
                destFile = new File(destPath);
            } else {
                destFile = new File(destPath);
                if(mac && destPath.indexOf("/")<0) destFile = new File(root, destPath);
                else if(!mac && destPath.indexOf("\\")<0) destFile = new File(root, destPath);
            }

            File dataFile = null;
            if(dataPath != null) {
                if(dataFolder != null) {
                    dataFile = new File(dataFolder, dataPath);
                } else {
                    dataFile = new File(dataPath);
                    if(mac && dataPath.indexOf("/")<0) dataFile = new File(root, dataPath);
                    else if(!mac && dataPath.indexOf("\\")<0) dataFile = new File(root, dataPath);
                }
            }
            ExcelTableInjector in = new ExcelTableInjector();
            if(command != null && command.compareToIgnoreCase("updateCell")==0) {
                in.updateCellValue(sourceFile.getAbsolutePath(), destFile.getAbsolutePath(), sheet, areaRef, value);
            } else if(command != null && command.compareToIgnoreCase("injectMultiple")==0) {
                in.doMultiple(batchInstructions, dataFolder, sourceFile.getAbsolutePath(), destFile.getAbsolutePath(), sheet);
            } else {
                in.excelTableInject(sourceFile.getAbsolutePath(), destFile.getAbsolutePath(), sheet, tableName, dataFile.getAbsolutePath());
            }
            System.out.println("OK.");

        } catch (Exception error) {
            ExcelTableInjector.printStackTrace(error);
            System.exit(-1);
        }
    }
 
}
