package excel;

import java.lang.reflect.Field;
import java.util.*;
import java.util.concurrent.Future;

/**
 * Created by chenchx on 2018/6/19.
 */
public class ReadExcel2007<T> extends ReadExcel2007Adapter {

    private Map<String,String> headMap = new HashMap<>();
    private Map<Integer,String> headNames = new HashMap<>();
    private int size = 100;
    private ReadExcel readExcel;
    private Class<T> tClass;
    private List<T> list = new ArrayList<>();

    public ReadExcel2007(Class<T> tClass, int size, ReadExcel readExcel){
        this.tClass = tClass;
        this.size = size;
        this.readExcel = readExcel;
        Field[] fields = tClass.getDeclaredFields();
        Arrays.stream(fields).forEach(field -> {
            ExcelHead excelHead = field.getAnnotation(ExcelHead.class);
            if (excelHead != null){
                headMap.put(excelHead.name(), field.getName());
            }
        });
    }

    @Override
    public void optRow(int curRow, List<String> rowList){
        Future future = null;
        if (curRow == 0){
            for (int i=0; i<rowList.size(); i++){
                headNames.put(i,rowList.get(i));
            }
        }else if (curRow == -1){
            if (list.size()>0){
                future = readExcel.optRows(list);
            }
        }else{
            try{
                T t = tClass.newInstance();
                for (int i=0;i<rowList.size();i++){
                    String titleName = headNames.get(i);
                    String fieldName = headMap.get(titleName);
                    String fieldValue = rowList.get(i);
                    if (fieldName != null){
                        reflectObject(t,fieldName,fieldValue);
                    }
                }
                list.add(t);
            }catch (Exception e){
                e.printStackTrace();
            }
            if (list.size() >= size){
                List<T> desc = new ArrayList<>();
                desc.addAll(desc);
                future = readExcel.optRows(desc);
                list = new ArrayList<>();
            }
        }
        if (future != null){
            if (futureList == null){
                futureList = new ArrayList<>();
            }
            futureList.add(future);
        }

    }

    private void reflectObject(T t, String fieldName, String fieldValue) throws NoSuchFieldException, IllegalAccessException {
        Field field = t.getClass().getDeclaredField(fieldName);
        field.setAccessible(true);
        field.set(t,fieldValue);
    }
}
