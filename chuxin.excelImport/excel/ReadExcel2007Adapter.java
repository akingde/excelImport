package excel;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.concurrent.Future;

/**
 * Created by chenchx on 2018/6/19.
 */
public abstract class ReadExcel2007Adapter extends DefaultHandler implements ImportExcel{

    protected List<Future> futureList;
    private XMLReader parser;
    private SharedStringsTable sst;
    private InputSource sheetSource;
    private InputStream sheet;
    private boolean nextIsString;
    private String lastContents;
    private List<String> rowlist = new ArrayList<>();
    //当前行
    private int curRow = 0;
    //当前列索引
    private int curCol = 0;
    //上一列列索引
    private int preCol = 0;
    //标题行，一般情况下为0
    private int titleRow = 0;
    //列数
    private int rowsize = 0;

    public abstract void optRow(int curRow, List<String> rowList);
    @Override
    public  List<Future> process() throws Exception {
        parser.parse(sheetSource);
        sheet.close();
        return futureList;
    }

    public ImportExcel init(InputStream inputStream) throws Exception{
        OPCPackage opcPackage = OPCPackage.open(inputStream);
        XSSFReader r = new XSSFReader(opcPackage);
        SharedStringsTable sst = r.getSharedStringsTable();
        parser = fetchSheetParser(sst);
        Iterator<InputStream> sheets = r.getSheetsData();
        sheet = sheets.next();
        sheetSource = new InputSource(sheet);
        return this;
    }

    private XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException {
        XMLReader parser = XMLReaderFactory.createXMLReader();
        this.sst = sst;
        parser.setContentHandler(this);
        return parser;
    }
    @Override
    public void startElement(String uri, String localName, String name, Attributes attributes){
        if (localName.equals("c")){
            String cellType = attributes.getValue("t");
            String rowStr = attributes.getValue("r");
            curCol = this.getRowIndex(rowStr);
            if (cellType != null && cellType.equals("s")) {

                nextIsString = true;
            } else {

                nextIsString = false;
            }
        }
        // 置空
        lastContents = "";
    }

    private int getRowIndex(String rowStr) {
        rowStr = rowStr.replaceAll("[^A-Z]", "");
        byte[] rowAbc = rowStr.getBytes();
        float num = 0;
        int len = rowAbc.length;
        for (int i=0;i<len;i++){
            num += (rowAbc[i]-'A'+1)*Math.pow(26,len-i-1 );
        }
        return (int) num;
    }
    @Override
    public void endElement(String uri, String localName, String name){
        if (nextIsString){
            try{
                int idx = Integer.parseInt(lastContents);
                lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
            }catch (Exception e){

            }

        }
        if (localName.equals("v")) {
            String value = lastContents.trim();
            value = value.equals("")?" ":value;
            int cols = curCol-preCol;
            if (cols>1) {
                for (int i = 0; i < cols - 1; i++) {
                    rowlist.add(preCol, "");
                }
            }
            preCol = curCol;
            rowlist.add(curCol-1, value);
        }else {
            //如果标签名称为 row ，这说明已到行尾，调用 optRows() 方法
            if (localName.equals("row")) {
                int tmpCols = rowlist.size();
                if(curRow>this.titleRow && tmpCols<this.rowsize){
                    for (int i = 0;i < this.rowsize-tmpCols;i++){
                        rowlist.add(rowlist.size(), "");
                    }
                }
                optRow(curRow,rowlist);
                if(curRow==this.titleRow){
                    this.rowsize = rowlist.size();
                }
                rowlist.clear();
                curRow++;
                curCol = 0;
                preCol = 0;
            }
        }
        if(localName.equals("worksheet")){
            optRow(-1,null);
        }
    }
    @Override
    public void characters(char[] ch, int start, int length){
        //得到单元格内容的值
        lastContents += new String(ch, start, length);
    }
}
