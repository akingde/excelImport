package excel;

import java.util.List;
import java.util.concurrent.Future;

/**
 * Created by chenchx on 2018/6/18.
 */
public interface ReadExcel {
    Future optRows(List dataList);
}
