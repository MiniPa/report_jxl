package chengjs.report_jxl;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;

import org.junit.After;
import org.junit.Assert;
import org.junit.Test;

import chengjs.report_jxl.consts.ConstsLocalTest;
import junit.framework.TestCase;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.WritableWorkbook;
import net.sf.json.JSONArray;
import net.sf.json.JSONObject;

public class ReprotExcelManagerTest extends TestCase{
  public ReprotExcelManager reprotExcelManager;
  public Workbook workbook;
  public JSONObject webParamsJsonObj;
  
  public void beforeAction() throws BiffException, IOException {
    reprotExcelManager = ReprotExcelManager.getReprotExcelManager();
    webParamsJsonObj = new JSONObject();
    webParamsJsonObj.put("exceltitle", "demoExcel");
    //1.读本地demoExcel.xlt
    workbook = reprotExcelManager.readModelWorkBook(webParamsJsonObj);
  }
  
  @Test public void testManage() throws Exception{
    beforeAction();
    JSONArray showOnWebJsonArray = new JSONArray();
    webParamsJsonObj.put("webParams1", "webParams1Value");
    OutputStream outputStream = new OutputStream() {
      @Override
      public void write(int b) throws IOException {
        // TODO Auto-generated method stub
      }
    };
    reprotExcelManager.manage(webParamsJsonObj, showOnWebJsonArray,outputStream,"export");
    System.out.println(showOnWebJsonArray);
  }
  
  @Test public void testReadModelWorkBook() throws BiffException, IOException {
    beforeAction();
    Assert.assertNotNull(workbook);
  }
 
  @Test public void testWriteReportWorkBook() throws Exception {
    beforeAction();
    File file = new File(System.getProperty("user.dir") + "\\" + ConstsLocalTest.LOCALFILE);
    WritableWorkbook writableWorkbook = Workbook.createWorkbook(file,workbook);
    writableWorkbook.write();
    writableWorkbook.close();
  }
  
  @After public void afterAction() {

  }
}










