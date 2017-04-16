package chengjs.report_jxl.consts;

import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * Excel 报表处理 常量类
 * @author Chengjs
 *
 */
public class ConstsReportExcel {
  
  public final static SimpleDateFormat DATE_FORMAT1 = new SimpleDateFormat("yyyy-MM-dd");
  
  /**
   * excel文件名
   */
  public final static String EXCELTITLE = "exceltitle";
  
  /**
   * excel 模板 相对项目根目录路径 
   */
  private static String EXCELMODELPATH = "";
  /**
   * excel 模板 相对项目根目录路径 
   */  
  public static String get_EXCELMODELPATH() {
    if (ConstsLocalTest.LOCALTEST) {
      EXCELMODELPATH  = "\\src\\main\\resources\\excelModel\\";
    } else {
      EXCELMODELPATH =  "\\excelmodel\\";
    }
    return EXCELMODELPATH;
  }
  
  /**
   * excel 模板 后缀名称 (97-2003excel版本) ".xlt"
   */
  public final static String EXCELAFFIX = ".xlt";
  
  /**
   * excel 写出报表 后缀名称 (97-2003excel版本)
   */
  public final static String REPORT_EXCEL_AFFIX =  DATE_FORMAT1.format(new Date()) + ".xls";
  
  /**
   * 接口实现类后缀规范
   */
  public final static String IMPL = "HandleImpl";
  
  /**
   * 接口实现类目录配置
   */
  private static String IMPLPATH = "";
  /**
   * 接口实现类目录配置
   */
  public static String get_IMPLPATH() {
    if (ConstsLocalTest.LOCALTEST) {
      IMPLPATH  = "chengjs.report_jxl.handle.";
    } else {
      IMPLPATH =  "cn.com.servyou.fcjyjkpt.app.excelreport.";
    }
    return IMPLPATH;
  }  
  
  /**
   * excel 模板中把参数格式 {$xxxx} 前缀
   */
  public final static String VAR_PREFIX = "\\$\\{";
  
  /**
   * excel 模板中把参数格式 {$xxxx} 后置
   */
  public final static String VAR_SUFFIX = "\\}";
  
  
  
}










