package chengjs.report_jxl.handle;

import java.util.HashMap;
import java.util.Iterator;
import java.util.Map.Entry;

import chengjs.report_jxl.consts.ConstsLocalTest;
import jxl.write.WritableSheet;
import net.sf.json.JSONArray;
import net.sf.json.JSONObject;

public class DemoExcelHandleImpl extends ProjectExcelHandle {

    private static final int COLSTART = 0;
    private static final int ROWSTART = 3;
    private static final int COLEND = 5;
    private static final int ROWEND = 7;
    
    //报表fields配置 
    private static final String[] FIELDS = {
        "col1","col2","col3","col4","col5","col6"
    };
  
  /**
   * 报表模块 实现类 继承ProjectExcelHandle 并实现calculateParamsValueMap方法
   * @param excelParamsMap excelModel中查找出来的需填充的参数
   * @param webParamsJsonObj 前段页面传入的报表相关参数
   * @param showOnWebJsonArray 传出到web用于构造grid报表展示的json
   * @return 
   */
  public void calculateParamsValueMap(HashMap<String, String> excelParamsSqlMap,
      JSONObject webParamsJsonObj, JSONArray showOnWebJsonArray) {
    // 1.根据构造好的默认 webGridJsonArray 更改参数 构造最终展示在页面上的json参数
    String param;
    String paramsSql;
    String paramValue;
    for (Entry<String, String> entry : excelParamsSqlMap.entrySet()) {
      param = entry.getKey();
      paramsSql = entry.getValue();
      paramValue = "";
      if (ConstsLocalTest.LOCALTEST) {
          paramValue = paramsSql + " Param Value";
          for (int i = 0; i < showOnWebJsonArray.size(); i++) {
            JSONObject jsonObject = (JSONObject)showOnWebJsonArray.get(i);
            Iterator<?> it = jsonObject.keys();
            while(it.hasNext()){    
              String key = (String) it.next();
              String value = jsonObject.getString(key);
              if(value.equals(param)){
                jsonObject.put(key, paramValue);
              }  
            }  
          }
      } else {
        // paramValue = getValueUseSql(paramsSql,webParamsJsonObj);
        // 1.更具不同的参数sql 匹配不同参数进行计算和查询
      }
      excelParamsValueMap.put(param, paramValue);
    }
  }

  /**
   * 实现构造默认的用于页面grid展示报表的jsonArray
   * 从writablesheet中根据grid起始的单元格序号实现
   * @param webParamsJsonObj
   * @return
   */
  @Override
  public void getDefaultShowOnWebGridJsonArray(WritableSheet writableSheet,
      JSONArray showOnWebJsonArray, JSONObject webParamsJsonObj) {
    // 1.根据 webjsonObj 构造单行空jsonobject
    // 2.根据起始 int colstart = "0", int rowstart= "2", int colend = "5", int rowend = "6" 
    //  遍历 构造初始的 showOnWebJsonArray
    JSONObject jsonObject;
    for (int i = ROWSTART; i <= ROWEND; i++) {
      jsonObject = new JSONObject();
      for (int j = COLSTART; j <= COLEND; j++) {
        String contents = writableSheet.getCell(j, i).getContents();
        jsonObject.put(FIELDS[j-COLSTART], contents);
      }
      showOnWebJsonArray.add(jsonObject);
    }
  }
  
  @Override
  public String getParamsSql(String params) {
    // demoExcel for test 
    return null;
  }
}










