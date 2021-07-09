package model;

import java.io.FileOutputStream;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Writer_Excel 
{
    private Workbook wb;
    private Sheet sheet;
    protected Row row;
    protected int row_insert;
    protected int total_cols;
    private String ext;

    // <editor-fold defaultstate="collapsed" desc="INITIAL">

    public Writer_Excel(boolean isSingleSheet)
    {
        init(isSingleSheet, "xlsx");
    }

    public Writer_Excel(boolean isSingleSheet, String ext)
    {
        init(isSingleSheet, ext);
    }

    private void init(boolean isSingleSheet, String ext)
    {
        set_output_type(ext);
        if(isSingleSheet)
            create_sheet("Sheet1");
    }

    private void set_output_type(String ext)
    {
        if(ext.equalsIgnoreCase("xls"))
            wb = new HSSFWorkbook();
        else if(ext.equalsIgnoreCase("xlsx"))
            wb = new XSSFWorkbook();
        this.ext = "." + ext;
    }

    protected void create_sheet(String sheet_name)
    {
        sheet = wb.createSheet(sheet_name);
    }

    // </editor-fold>

    // <editor-fold defaultstate="collapsed" desc="ADD NEW DATASET NEXT TO THE LAST ROW">

    protected void add_rowData_nextToLastRow(String[] arr_data)
    {
        create_new_row();
        Writer_Excel.this.set_row_data(row, arr_data);
    }

    protected void add_rowData_nextToLastRow(List<String> dataList)
    {
        create_new_row();
        Writer_Excel.this.set_row_data(row, dataList);
    }

    protected void add_rowData_nextToLastRow(Cell[] arr_data)
    {
        create_new_row();
        Writer_Excel.this.set_row_data(row, arr_data);
    }
    
    private void create_new_row()
    {
        row_insert = (sheet.getLastRowNum() > 0) ? sheet.getLastRowNum() + 1 : 1;
        row = sheet.createRow(row_insert);
    }
    
    // </editor-fold>

    // <editor-fold defaultstate="collapsed" desc="ADD NEW DATASET NEXT TO THE LAST COLUMN OF CURRENT ROW">

    protected void add_colData_num_nextToLastCol(List<Double> dataList)
    {
        set_row_data_num(row, row.getLastCellNum(), dataList);
    }
    
    protected void add_colData_nextToLastCol(List<String> dataList)
    {
        set_row_data(row, row.getLastCellNum(), dataList);
    }

    // </editor-fold>

    // <editor-fold defaultstate="collapsed" desc="SET NEW DATA TO SPECIFIC ROW">

    protected void set_row_data(Row row, int col_i, String data)
    {
        row.createCell(col_i).setCellValue(data);
    }

    protected void set_row_data(Row row, int col_i, double data)
    {
        row.createCell(col_i).setCellValue(data);
    }
    
    // </editor-fold>

    // <editor-fold defaultstate="collapsed" desc="SET NEW DATASET TO SPECIFIC ROW">

    protected void set_row_data(Row row, double[] arr_data)
    {
        total_cols = arr_data.length;
        for(int i = 0; i < total_cols; i++)
            row.createCell(i).setCellValue(arr_data[i]);
    }
    
    protected void set_row_data(Row row, String[] arr_data)
    {
        total_cols = arr_data.length;
        for(int i = 0; i < total_cols; i++)
            row.createCell(i).setCellValue(arr_data[i]);
    }

    protected void set_row_data(Row row, Cell[] arr_data)
    {
        total_cols = arr_data.length;
        for(int i = 0; i < total_cols; i++)
            row.createCell(i).setCellValue(arr_data[i].getStringCellValue());
    }

    protected void set_row_data(Row row, List<String> dataList)
    {
        total_cols = dataList.size();
        for(int i = 0; i < total_cols; i++)
            row.createCell(i).setCellValue(dataList.get(i));
    }
    
    protected void set_row_data_num(Row row, List<Double> dataList)
    {
        total_cols = dataList.size();
        for(int i = 0; i < total_cols; i++)
            row.createCell(i).setCellValue(dataList.get(i));
    }

    // </editor-fold>

    // <editor-fold defaultstate="collapsed" desc="SET NEW DATASET TO SPECIFIC ROW WITH SPECIFIC COLUMN INDEX TO START TO SET">

    protected void set_row_data(Row row, int col_i_begin, double[] arr_data)
    {
        total_cols = arr_data.length;
        for(int i = 0; i < total_cols; i++)
            row.createCell(col_i_begin + i).setCellValue(arr_data[i]);
    }

    protected void set_row_data(Row row, int col_i_begin, String[] arr_data)
    {
        total_cols = arr_data.length;
        for(int i = 0; i < total_cols; i++)
            row.createCell(col_i_begin + i).setCellValue(arr_data[i]);
    }

    protected void set_row_data(Row row, int col_i_begin, Cell[] arr_data)
    {
        total_cols = arr_data.length;
        for(int i = 0; i < total_cols; i++)
            row.createCell(col_i_begin + i).setCellValue(arr_data[i].getStringCellValue());
    }

    protected void set_row_data(Row row, int col_i_begin, List<String> dataList)
    {
        total_cols = dataList.size();
        for(int i = 0; i < total_cols; i++)
            row.createCell(col_i_begin + i).setCellValue(dataList.get(i));
    }

    protected void set_row_data_num(Row row, int col_i_begin, List<Double> dataList)
    {
        total_cols = dataList.size();
        for(int i = 0; i < total_cols; i++)
            row.createCell(col_i_begin + i).setCellValue(dataList.get(i));
    }

    // </editor-fold>

    public boolean is_row_existed()
    {
        return sheet.getLastRowNum() > 0;
    }

    public void export(String path2output) throws Exception
    {
        alertDefaultMsg(1);
        FileOutputStream output = new FileOutputStream(path2output + ext);
        wb.write(output);
        output.close();
        wb.close();
        alertDefaultMsg(3);
    }

    public void alertDefaultMsg(int msg_i)
    {
        switch(msg_i)
        {
            case 1: System.out.print("\r\tExporting data to excel file"); break;
            case 2: System.err.println("\tERROR to convert to excel file"); break;
            case 3: System.out.println("\r\tConvert to excel file successful."); break;
        }
    }
}