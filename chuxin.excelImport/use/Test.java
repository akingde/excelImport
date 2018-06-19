package use;

import excel.ImportUtil;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.List;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.Future;

/**
 * Created by chenchx on 2018/6/19.
 */
public class Test{
    public void handlerExcel(File file){
        InputStream in = null;
        try {
            in = new FileInputStream(file);
            List<Future> futureList = ImportUtil.getImportExcel(file.getName(),in,Test.class ,1000, dataList ->{
                Future future = CompletableFuture.supplyAsync(()->{
                    doSomeThing(dataList);
                    return null;
                });
                return future;
            }).process();
            for (Future future : futureList) {
                future.get();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    private void doSomeThing(List dataList) {
    }

}
