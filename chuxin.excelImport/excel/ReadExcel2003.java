package excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.*;
import java.util.concurrent.Future;

/**
 * Created by chenchx on 2018/6/18.
 */
public class ReadExcel2003<T> implements ImportExcel{

    private Class<T> tClass;
    private int size = 100;
    private Sheet sheet;
    private ReadExcel readExcel;
    private Map<String,String> headMap = new HashMap<>();
    private List<Future> futureList;


    public ReadExcel2003(Class<T> tClass, InputStream inputStream, int size, ReadExcel readExcel){
        this.tClass = tClass;
        this.readExcel = readExcel;
        this.size = size;
        try {
            this.sheet = new HSSFWorkbook(inputStream).getSheetAt(0);
        } catch (IOException e) {
            e.printStackTrace();
        }
        Field[] fields = tClass.getDeclaredFields();
        Arrays.stream(fields).forEach(field -> {
                ExcelHead excelHead = field.getAnnotation(ExcelHead.class);
                if (null != excelHead){
                    this.headMap.put(excelHead.name(), field.getName());
                }

            }
        );
    }

    @Override
    public List<Future> process() throws Exception {
        Row headRow = sheet.getRow(0);
        Map<Integer,String> headNames = doWithRow(headRow);
        int total = sheet.getLastRowNum();
        int pages = total / size + (total % size == 0 ? 0 : 1);
        for (int page = 0; page<pages; page++){
            List<T> tList = new ArrayList<>();
            for (int rowIndex = page * size + 1; rowIndex <= size * (page + 1) && rowIndex <= total; rowIndex++){
                Row row = sheet.getRow(rowIndex);
                Map<Integer,String> values = doWithRow(row);
                T t = tClass.newInstance();
                for (Integer colIndex : headNames.keySet()){
                    String headName = headNames.get(colIndex);
                    String value = values.get(colIndex);
                    String fieldName = headMap.get(headName);
                    if (null != fieldName){
                        reflectObject(t, fieldName, value);
                    }
                }
                tList.add(t);
            }
            if (tList.size()>0){
                Future future = readExcel.optRows(tList);
                if (futureList == null){
                    futureList = new ArrayList<>();
                }
                futureList.add(future);
            }
        }
        return futureList;
    }

    private void reflectObject(T t, String fieldName, String value) throws IllegalAccessException, NoSuchFieldException {
        Field field = t.getClass().getDeclaredField(fieldName);
        field.setAccessible(true);
        field.set(t,value);
    }

    private Map<Integer,String> doWithRow(Row headRow) {
        Map<Integer,String> rowMap = new HashMap<>(16);
        Iterator<Cell> iterator = headRow.iterator();
        while (iterator.hasNext()){
            Cell cell = iterator.next();
            cell.setCellType(CellType.STRING);
            String value = cell.getStringCellValue();
            rowMap.put(cell.getColumnIndex(),value);
        }
        return rowMap;
    }
}
