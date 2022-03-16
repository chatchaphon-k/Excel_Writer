package extensions;

import java.awt.Desktop;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public abstract class Writer_Excel 
{
    private String ext;
    private Workbook wb;
    private Sheet sheet;
    protected Row row;
    protected int row_insert;
    protected int total_cols;

    // <editor-fold defaultstate="collapsed" desc="INITIAL">

    public Writer_Excel(boolean isSingleSheet)
    {
        init(isSingleSheet, "xlsx");
    }

    public Writer_Excel(boolean isSingleSheet, String ext)
    {
        init(isSingleSheet, ext);
    }

    public void init(boolean isSingleSheet, String ext)
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
        set_row_data(row, arr_data);
    }

    protected void add_rowData_nextToLastRow(List<String> dataList)
    {
        create_new_row();
        set_row_data(row, dataList);
    }

    protected void add_rowData_nextToLastRow(Cell[] arr_data)
    {
        create_new_row();
        set_row_data(row, arr_data);
    }

    protected void add_rowData_nextToLastRow(double[] arr_data)
    {
        create_new_row();
        set_row_data_num(row, arr_data);
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

    protected void set_row_data_num(Row row, double[] dataList)
    {
        total_cols = dataList.length;
        for(int i = 0; i < total_cols; i++)
            row.createCell(i).setCellValue(dataList[i]);
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

    // <editor-fold defaultstate="collapsed" desc="GETTER">

    public Workbook wb()
    {
        return wb;
    }

    public Sheet sheet()
    {
        return sheet;
    }

    // </editor-fold>

    // <editor-fold defaultstate="collapsed" desc="SETTER">

    protected void set_sheet(String sheet_name)
    {
        sheet = wb.getSheet(sheet_name);
    }
    
    protected void set_column_name(int row_i, String[] cols_name)
    {
        row = sheet().createRow(row_i);
        set_row_data(row, cols_name);
    }

    protected void set_column_name(int row_i, List<String> cols_name)
    {
        row = sheet().createRow(row_i);
        set_row_data(row, cols_name);
    }

    protected void set_column_name(String[] cols_name)
    {
        set_column_name(0, cols_name);
    }

    protected void set_column_name(List<String> cols_name)
    {
        set_column_name(0, cols_name);
    }

    // </editor-fold>

    public boolean is_row_existed()
    {
        return sheet.getLastRowNum() > 0;
    }

    public void export_wOpenDir(String path2resDir, String output_filename) throws Exception
    {
        export(path2resDir, output_filename);
        open_dir(path2resDir);
    }

    public void export(String path2resDir, String output_filename) throws Exception
    {
        export(path2resDir + "\\" + output_filename);
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
            case 1: System.out.print("\r\t\t- Exporting data to excel file"); break;
            case 2: System.err.println("\t\t- ERROR to convert to excel file"); break;
            case 3: System.out.println("\r\t\t- Convert to excel file successful"); break;
        }
    }

    public void open_dir(String path2resDir) throws IOException
    {
        Desktop.getDesktop().open(new File(path2resDir));
    }
}