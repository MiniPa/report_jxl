package chengjs.report_jxl;

import chengjs.report_jxl.consts.ConstsReportExcel;
import chengjs.report_jxl.handle.IExcelHandle;
import chengjs.report_jxl.util.UtilString;

/**
 * 利用反射原理 根据 model 名称来创建类
 * @author Chengjs
 */
public class ReportExcelFactory {

  private static ReportExcelFactory uniqueInstance = null;
  
  /**
   * 反射构造 excel报表 处理类 IExcelHandle
   * @param model 模板逻辑处理类名称中 模板名称部分 从web传入
   * @return
   * @throws ClassNotFoundException
   * @throws InstantiationException
   * @throws IllegalAccessException
   */
  public static IExcelHandle create(String model) throws Exception {
    String implClass = ConstsReportExcel.get_IMPLPATH() 
        + UtilString.toUpperCaseFirstOne(model) 
        + ConstsReportExcel.IMPL;
    Class<?> cls = java.lang.Class.forName(implClass);
    IExcelHandle handle = (IExcelHandle) cls.newInstance();
    return handle;
  }
  
  public static ReportExcelFactory getReportExcelFactory(){
    if(uniqueInstance == null){
      synchronized(ReportExcelFactory.class){
        if(uniqueInstance == null){
          uniqueInstance=new ReportExcelFactory();
        }
      }
    }
    return uniqueInstance;
  }
  
  public ReportExcelFactory() {
    super();
  }
  
  
}










