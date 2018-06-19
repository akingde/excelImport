package excel;

import java.io.InputStream;

/**
 * Created by chenchx on 2018/6/19.
 */
public class ImportUtil {

    /**
     *
     * @param fileName 文件名 用于判断excel版本
     * @param in 文件名 用于判断excel版本
     * @param tClass 返回结果类
     * @param pageSize 一次读取数量
     * @param readExcel 一次读取的结果处理
     *
     *
     **/
    public static <T> ImportExcel getImportExcel(String fileName, InputStream in, Class<T> tClass, int pageSize, ReadExcel readExcel){
        if(fileName.endsWith(".xlsx")){
            //2007
            try {
                return new ReadExcel2007(tClass, pageSize,readExcel).init(in);
            } catch (Exception e) {
                e.printStackTrace();
            }
        }else {
            //2003
            return new ReadExcel2003(tClass, in, pageSize,readExcel);

        }
        return null;
    }
}
