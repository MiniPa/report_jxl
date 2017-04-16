package chengjs.report_jxl;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;

import chengjs.report_jxl.consts.ConstsLocalTest;
import chengjs.report_jxl.consts.ConstsReportExcel;
import chengjs.report_jxl.handle.IExcelHandle;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import net.sf.json.JSONArray;
import net.sf.json.JSONObject;

/**
 * excel报表管理类
 * 读取、修改、写出excel 报表
 * @author Chengjs
 *
 */
public class ReprotExcelManager {
  
  private static ReprotExcelManager uniqueInstance = null;
  
  /**
   * 管理excel报表生成整个流程 流程入口
   * @param webParamsJsonObj
   * @param OutputStream out WritableWorkBook 输出流
   * @param String flag 为"query"--不使用输出流, 或 "export"--导出 使用输出流
   * @return
   * @throws Exception 
   */
  public WritableWorkbook manageGrid(JSONObject webParamsJsonObj,JSONArray showOnWebJsonArray) throws Exception {
    Workbook workbook = readModelWorkBook(webParamsJsonObj);// 读
    String modelTitle = webParamsJsonObj.getString("exceltitle");
    File file;
    WritableWorkbook writableWorkbook;
    if (ConstsLocalTest.LOCALTEST) {
      file = new File(System.getProperty("user.dir") + "\\" + ConstsLocalTest.LOCALFILE);
    } else {
      file = File.createTempFile(modelTitle, ConstsReportExcel.REPORT_EXCEL_AFFIX); 
    }
    writableWorkbook = Workbook.createWorkbook(file,workbook);
    
    //用reportExcelFacotry 创建excelhandle对象 进行相应的excel处理
    IExcelHandle handle = ReportExcelFactory.create(modelTitle);
    
    modifyReportWorkBook(writableWorkbook,handle,webParamsJsonObj,showOnWebJsonArray);
//    if (ConstsLocalTest.LOCALTEST) {
//      try {
//        writableWorkbook.write();
//        writableWorkbook.close();
//      } catch (WriteException e) {
//        e.printStackTrace();
//      }
//    }
    return writableWorkbook;
  }
  
  /**
   * 管理excel报表生成整个流程 流程入口
   * @param webParamsJsonObj
   * @param OutputStream out WritableWorkBook 输出流
   * @param String flag 为"query"--不使用输出流, 或 "export"--导出 使用输出流
   * @return
   * @throws Exception 
   */
  public WritableWorkbook manage(JSONObject webParamsJsonObj,JSONArray showOnWebJsonArray,OutputStream out,String flag) throws Exception {
    Workbook workbook = readModelWorkBook(webParamsJsonObj);// 读
    String modelTitle = webParamsJsonObj.getString("exceltitle");
    File file;
    WritableWorkbook writableWorkbook;
    if (ConstsLocalTest.LOCALTEST) {
      file = new File(System.getProperty("user.dir") + "\\" + ConstsLocalTest.LOCALFILE);
    } else {
      file = File.createTempFile(modelTitle, ConstsReportExcel.REPORT_EXCEL_AFFIX); 
    }
    if (ConstsLocalTest.WWB_LOCAL || "query".equals(flag) ) {
      writableWorkbook = Workbook.createWorkbook(file,workbook);
    } else {
        writableWorkbook = Workbook.createWorkbook(out, workbook);
    }
    
    //用reportExcelFacotry 创建excelhandle对象 进行相应的excel处理
    IExcelHandle handle;
    handle = ReportExcelFactory.create(modelTitle);
    
    modifyReportWorkBook(writableWorkbook,handle,webParamsJsonObj,showOnWebJsonArray);
    if (ConstsLocalTest.LOCALTEST) {
//      try {
//        writableWorkbook.write();
//        writableWorkbook.close();
//      } catch (WriteException e) {
//        e.printStackTrace();
//      }
    }
    return writableWorkbook;
  }
  
  /**
   * 读取本地excel报表模板
   * @param webParamsJsonObj 浏览器传入的报表参数信息
   * @return
   * @throws BiffException
   * @throws IOException
   */
  public Workbook readModelWorkBook(JSONObject webParamsJsonObj) throws BiffException, IOException {
    String workbookTitle = webParamsJsonObj.getString(ConstsReportExcel.EXCELTITLE);
    String relativelyPath=System.getProperty("user.dir"); 
    String excelPath = relativelyPath + ConstsReportExcel.get_EXCELMODELPATH() + workbookTitle + ConstsReportExcel.EXCELAFFIX;
    File file = new File(excelPath);
    Workbook workbook = Workbook.getWorkbook(file);
    return workbook;
  }
  
  /**
   * 修改sheet相应内容
   * @param writableWorkbook
   * @param handle
   * @param webParamsJsonObj
   */
  private void modifyReportWorkBook(WritableWorkbook writableWorkbook, 
      IExcelHandle handle,JSONObject webParamsJsonObj, JSONArray showOnWebJsonArray) {
    WritableSheet writableSheet = writableWorkbook.getSheet(0);
    handle.getDefaultShowOnWebGridJsonArray(writableSheet,showOnWebJsonArray,webParamsJsonObj);
    handle.handle(writableSheet, webParamsJsonObj, showOnWebJsonArray);
  }
  
  /**
   * 写workbook
   * @param workbook
   * @param webParamsJsonObj
   * @throws IOException
   * @throws WriteException
   */
  public void writeReportWorkBook(WritableWorkbook writableWorkbook) 
       throws IOException, WriteException {
    writableWorkbook.write();
    writableWorkbook.close();
  }

  public static ReprotExcelManager getReprotExcelManager(){
    if(uniqueInstance == null){
      synchronized(ReprotExcelManager.class){
        if(uniqueInstance == null){
          uniqueInstance=new ReprotExcelManager();
        }
      }
    }
    return uniqueInstance;
  }
  
  public ReprotExcelManager() {
    super();
  }
}










