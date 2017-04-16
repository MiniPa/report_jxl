package chengjs.report_jxl.handle;

import java.util.HashMap;
import java.util.Map.Entry;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import chengjs.report_jxl.consts.ConstsLocalTest;
import chengjs.report_jxl.consts.ConstsReportExcel;
import jxl.Cell;
import jxl.LabelCell;
import jxl.Sheet;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;
import net.sf.json.JSONArray;
import net.sf.json.JSONObject;

public abstract class ProjectExcelHandle implements IExcelHandle {
  
  /**
   * 报表参数值集合
   */
  protected HashMap<String, String> excelParamsValueMap = new HashMap<String, String>();
  
  /**
   * 引用此包需重写并继承此类 作为集成 IHandle的父类 并实现此方法
   * @param param
   * @return
   */
  public abstract String getParamsSql(String params);

  /**
   * 处理WritableSheet
   * @param writableSheet WritableSheet
   * @param webParamsJsonObj web传入的报表条件参数
   */
  @Override
  public void handle(WritableSheet writableSheet, JSONObject webParamsJsonObj, JSONArray showOnWebJsonArray) {
    // 1.获取sheet中需要替换的参数集合 params--Cell
    HashMap<String, Cell> excelParamsTitleMap = getExecelParamsMap(writableSheet);
    // 2.获取参数相应取报表值的sql params--Sql
    HashMap<String, String> excelParamsSqlMap = new HashMap<String, String>();
    excelParamsSqlMap = getParamsSql(excelParamsTitleMap, excelParamsSqlMap);
    // 3.计算参数对应的各项实际值 excelParamsValueMap params--trueValue
    //    并构造传递到web的参数
    calculateParamsValueMap(excelParamsSqlMap, webParamsJsonObj, showOnWebJsonArray);
    // 4.将参数各项值放入报表
    flushValueToSheet(writableSheet, excelParamsTitleMap);
  }  
  
  /**
   * 将参数各项值放入报表
   * @param writableSheet
   * @param excelParamsTitleMap
   */
  protected void flushValueToSheet(WritableSheet writableSheet, 
      HashMap<String, Cell> excelParamsTitleMap) {
    String param = "";
    try {
      for (Entry<String, Cell> entry : excelParamsTitleMap.entrySet()) {
        param = entry.getKey();
        LabelCell labelCell = writableSheet.findLabelCell(param);
        Label label = new Label(labelCell);
        label.setString(excelParamsValueMap.get(param));
          writableSheet.addCell(label);
      }
    } catch (RowsExceededException e) {
      e.printStackTrace();
    } catch (WriteException e) {
      e.printStackTrace();
    }
  }  
  
  /**
   * 根据条件 查找查询参数对应的sql
   * @param excelParamsTitleMap params-Cell k-v
   * @param excelParamsSqlMap parmas-Sql k-v
   * @return HashMap<String, String> 对应参数的查询 params-sql
   */
  protected HashMap<String, String> getParamsSql(HashMap<String, Cell> excelParamsTitleMap, 
      HashMap<String, String> excelParamsSqlMap) {
    
    String paramsSql;
    Set<String> paramsSet = excelParamsTitleMap.keySet();
    for (String param : paramsSet) {
      paramsSql = "";
      if (ConstsLocalTest.LOCALTEST) {
          paramsSql = param +" testsql";
      } else {
        paramsSql = getParamsSql(param);
      }
      excelParamsSqlMap.put(param, paramsSql);
    }
    return excelParamsSqlMap;
  };
  
  /**
   * 获取sheet中符合参数样式的单元格内容和单元格对象
   * @param sheet
   * @return map String--excel模板参数内容, Cell含参数标记的单元格
   */
  public HashMap<String, Cell> getExecelParamsMap(Sheet sheet) {
    String regex = ConstsReportExcel.VAR_PREFIX + "(\\w)+(\\-)?(\\w)+" + ConstsReportExcel.VAR_SUFFIX;
    HashMap<String, Cell> excelParamsTitleMap = new HashMap<String, Cell>();
    //获得指定Sheet含有的行数
    int num = sheet.getRows();
    //循环读取数据
    for(int i=0;i<num;i++) {
        //得到第i行的数据..返回cell数组
        Cell[] cell = sheet.getRow(i);
        //装载读取数据的集合
        for(int j = 0; j < cell.length; j++) {
          String paramValue = extract(regex, cell[j].getContents());
          if (!"".equals(paramValue)) {
            excelParamsTitleMap.put(paramValue, cell[j]);
          }
        }
    }
    return excelParamsTitleMap;
  }
  
  /** 
   * 萃取字符串
   * @param regex
   * @param input
   * @return
   */
  protected String extract(String regex, String input){
      String ret = "";
      Pattern r = Pattern.compile(regex);
      Matcher m = r.matcher(input);
      if (m.find())
          ret = m.group();
      return ret;
  }
  
}










