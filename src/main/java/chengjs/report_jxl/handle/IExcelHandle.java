package chengjs.report_jxl.handle;

import java.util.HashMap;

import jxl.write.WritableSheet;
import net.sf.json.JSONArray;
import net.sf.json.JSONObject;

/**
 * Excel Sheet 处理类接口
 * @author Chengjs
 */
public interface IExcelHandle {

  /**
   * 调用此方法来执行报表参数处理流程
   * 处理WritableSheet
   * @param writableSheet WritableSheet
   * @param webParamsJsonObj web传入的报表条件参数
   * @param showOnWebJsonArray 传出到web用于构造grid报表展示的json
   */
  void handle(WritableSheet writableSheet, JSONObject webParamsJsonObj, JSONArray showOnWebJsonArray);
  
  /**
   * 实现类需自定义实现此方法
   * @param excelParamsMap excelModel中查找出来的需填充的参数
   * @param webParamsJsonObj 前段页面传入的报表相关参数
   * @param showOnWebJsonArray 传出到web用于构造grid报表展示的json
   * @return 
   */
  void calculateParamsValueMap(HashMap<String, String> excelParamsSqlMap,
      JSONObject webParamsJsonObj,JSONArray showOnWebJsonArray);
  
  /**
   * 实现构造默认的用于页面grid展示报表的jsonArray
   * 从writablesheet中根据grid起始的单元格序号实现
   * @param showOnWebJsonArray
   * @return
   */
  void getDefaultShowOnWebGridJsonArray(WritableSheet writableSheet, JSONArray showOnWebJsonArray, JSONObject webParamsJsonObj);
  
}










